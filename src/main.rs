use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use std::sync::atomic::{AtomicUsize, AtomicBool, Ordering};
use std::time::{Duration, Instant};
use std::thread;
use windows::{
    core::*,
    Win32::{
        Foundation::*,
        System::LibraryLoader::*,
        System::Threading::*,
        System::Console::*,
        UI::WindowsAndMessaging::*,
        UI::Input::KeyboardAndMouse::*,
    },
};

// Define WM constants
const WM_USER: u32 = 0x0400;
const WM_KEYDOWN: u32 = 0x0100;
const WM_SYSKEYDOWN: u32 = 0x0104;

// Custom messages for deferred operations
const WM_TILER_COMMAND: u32 = WM_USER + 2;
const WM_TILER_SHUTDOWN: u32 = WM_USER + 3;
const WM_TILER_RECALC: u32 = WM_USER + 4;

// Command types for deferred execution
#[derive(Debug, Clone, Copy)]
#[repr(u32)]
enum TilerCommand {
    PanLeft = 0,
    PanRight = 1,
    PanUp = 2,
    PanDown = 3,
    ResizeLeft = 4,
    ResizeRight = 5,
    MoveUp = 6,
    MoveDown = 7,
    MoveLeft = 8,
    MoveRight = 9,
    AddWindow = 10,
    IncreaseTransparency = 14,
    DecreaseTransparency = 15,
    ScrollToWindow = 17,
    IncreaseMargins = 18,
    DecreaseMargins = 19,
    RemoveWindow = 20,
    CycleFPS = 21,
    ForceRecalc = 22,
}

// Animation state for smooth transitions
#[derive(Debug, Clone)]
struct AnimationState {
    start_rect: RECT,
    target_rect: RECT,
    start_time: Instant,
    duration: Duration,
    animation_type: AnimationType,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum AnimationType {
    Move,       // Normal movement animation
    Entry,      // Scale up from center on entry
    Exit,       // Scale down to center on exit
}

// Combined scroll animation state
#[derive(Debug, Clone)]
struct ScrollAnimation {
    start_x: i32,
    start_y: i32,
    target_x: i32,
    target_y: i32,
    start_time: Instant,
    duration: Duration,
}

// Window size variants - simplified to just width variations
#[derive(Debug, Clone, Copy, PartialEq)]
enum TileSize {
    Full,           // Full screen width
    Half,           // Half screen width
}

// Position in the ribbon (x is the virtual position, row is the vertical row)
#[derive(Debug, Clone, Copy)]
struct RibbonPosition {
    x: i32,         // Virtual x position in ribbon
    row: i32,       // Row number (0, 1, 2, etc.)
    size: TileSize,
}

// Managed window information
#[derive(Debug, Clone)]
struct ManagedWindow {
    hwnd: HWND,
    original_style: WINDOW_STYLE,
    original_ex_style: WINDOW_EX_STYLE,
    original_rect: RECT,
    position: RibbonPosition,
    animation: Option<AnimationState>,
}

// Command queue entry
struct QueuedCommand {
    command: TilerCommand,
    hwnd: HWND,
    timestamp: Instant,
}

// Main tiler state
struct RibbonTiler {
    windows: HashMap<isize, ManagedWindow>,
    floating_windows: HashMap<isize, HWND>,
    ribbon_offset: i32,
    ribbon_offset_target: i32,
    vertical_offset: i32,              // Current vertical scroll offset
    vertical_offset_target: i32,       // Target vertical scroll offset
    scroll_animation: Option<ScrollAnimation>, // Combined scroll animation
    current_row: i32,                  // Currently visible row
    row_height: i32,                   // Height of each row
    monitor_width: i32,
    monitor_height: i32,
    last_resolution_check: Instant,
    resolution_check_throttle_ms: u64,
    margin_horizontal: i32,
    margin_vertical: i32,
    transparency: u8,
    animation_running: Arc<Mutex<bool>>,
    animation_stop_requested: Arc<Mutex<bool>>,
    main_thread_id: u32,
    main_hwnd: HWND,
    command_queue: Vec<QueuedCommand>,
    last_command_time: HashMap<u32, Instant>,
    animation_fps: u64,
    needs_ribbon_recalc: bool,
    last_ribbon_recalc: Instant,
}

impl RibbonTiler {
    fn new() -> Self {
        let (width, height) = Self::get_monitor_dimensions();
        let main_thread_id = unsafe { GetCurrentThreadId() };
        
        // Create a hidden window for message processing
        let main_hwnd = unsafe {
            let hwnd = CreateWindowExW(
                WINDOW_EX_STYLE(0),
                w!("STATIC"),
                w!("RibbonTilerMessageWindow"),
                WS_OVERLAPPED,
                0, 0, 0, 0,
                HWND::default(),
                HMENU::default(),
                GetModuleHandleW(None).unwrap_or_default(),
                None
            );
            
            if hwnd.0 == 0 {
                panic!("Failed to create message window");
            }
            hwnd
        };
        
        Self {
            windows: HashMap::new(),
            floating_windows: HashMap::new(),
            ribbon_offset: 0,
            ribbon_offset_target: 0,
            vertical_offset: 0,
            vertical_offset_target: 0,
            scroll_animation: None,
            current_row: 0,
            row_height: height,  // Each row is full monitor height
            monitor_width: width,
            monitor_height: height,
            last_resolution_check: Instant::now(),
            resolution_check_throttle_ms: 1000,
            margin_horizontal: 40,
            margin_vertical: 80,
            transparency: 255,
            animation_running: Arc::new(Mutex::new(false)),
            animation_stop_requested: Arc::new(Mutex::new(false)),
            main_thread_id,
            main_hwnd,
            command_queue: Vec::new(),
            last_command_time: HashMap::new(),
            animation_fps: 90,
            needs_ribbon_recalc: false,
            last_ribbon_recalc: Instant::now(),
        }
    }

    fn get_monitor_dimensions() -> (i32, i32) {
        unsafe {
            let screen_width = GetSystemMetrics(SM_CXSCREEN);
            let screen_height = GetSystemMetrics(SM_CYSCREEN);
            (screen_width, screen_height)
        }
    }

    // Process queued commands - called from message loop
    fn process_command_queue(&mut self) {
        let commands = std::mem::take(&mut self.command_queue);
        let now = Instant::now();
        
        // Clean up closed windows before processing commands
        self.clean_closed_windows();
        
        for queued in commands {
            let should_throttle = match queued.command {
                TilerCommand::PanLeft | TilerCommand::PanRight | TilerCommand::PanUp | TilerCommand::PanDown => false,
                _ => true,
            };
            
            if should_throttle {
                if let Some(&last_time) = self.last_command_time.get(&(queued.command as u32)) {
                    if now.duration_since(last_time).as_millis() < 50 {
                        continue;
                    }
                }
            }
            
            self.last_command_time.insert(queued.command as u32, now);
            
            match queued.command {
                TilerCommand::PanLeft => self.pan_ribbon(Direction::Left),
                TilerCommand::PanRight => self.pan_ribbon(Direction::Right),
                TilerCommand::PanUp => self.pan_row(Direction::Up),
                TilerCommand::PanDown => self.pan_row(Direction::Down),
                TilerCommand::ResizeLeft => self.resize_window(queued.hwnd, Direction::Left),
                TilerCommand::ResizeRight => self.resize_window(queued.hwnd, Direction::Right),
                TilerCommand::MoveUp => self.move_window(queued.hwnd, Direction::Up),
                TilerCommand::MoveDown => self.move_window(queued.hwnd, Direction::Down),
                TilerCommand::MoveLeft => self.move_window(queued.hwnd, Direction::Left),
                TilerCommand::MoveRight => self.move_window(queued.hwnd, Direction::Right),
                TilerCommand::AddWindow => { self.add_window(queued.hwnd); },
                TilerCommand::IncreaseTransparency => self.adjust_transparency(10),
                TilerCommand::DecreaseTransparency => self.adjust_transparency(-10),
                TilerCommand::ScrollToWindow => {
                    if self.windows.contains_key(&queued.hwnd.0) {
                        self.scroll_to_window(queued.hwnd);
                    }
                },
                TilerCommand::IncreaseMargins => self.adjust_margins(5),
                TilerCommand::DecreaseMargins => self.adjust_margins(-5),
                TilerCommand::RemoveWindow => {
                    self.remove_window(queued.hwnd);
                    if self.needs_ribbon_recalc {
                        self.recalculate_ribbon();
                    }
                },
                TilerCommand::CycleFPS => self.cycle_fps(),
                TilerCommand::ForceRecalc => {
                    self.clean_closed_windows();
                    self.recalculate_ribbon();
                },
            }
        }
    }

    fn queue_command(&mut self, command: TilerCommand, hwnd: HWND) {
        self.command_queue.push(QueuedCommand {
            command,
            hwnd,
            timestamp: Instant::now(),
        });
    }

    fn ease_out_cubic(t: f32) -> f32 {
        let t = t - 1.0;
        t * t * t + 1.0
    }

    fn lerp(start: i32, end: i32, t: f32) -> i32 {
        start + ((end - start) as f32 * t) as i32
    }

    fn is_window_visible(&self, pos: &RibbonPosition) -> bool {
        // Check horizontal visibility
        let window_start = pos.x - self.ribbon_offset;
        let window_end = window_start + self.get_tile_width(&pos.size);
        let h_visible = window_end >= -self.monitor_width && window_start <= self.monitor_width * 2;
        
        // Check vertical visibility
        let window_top = pos.row * self.row_height - self.vertical_offset;
        let window_bottom = window_top + self.row_height;
        let v_visible = window_bottom >= -self.row_height && window_top <= self.monitor_height + self.row_height;
        
        h_visible && v_visible
    }

    // Update animations
    fn update_animations(&mut self) {
        let now = Instant::now();
        let mut animations_complete = Vec::new();
        let mut window_updates = Vec::new();
        let mut need_reposition = false;

        // Update combined scroll animation
        if let Some(scroll_anim) = &self.scroll_animation {
            let elapsed = now.duration_since(scroll_anim.start_time);
            
            if elapsed >= scroll_anim.duration {
                self.ribbon_offset = scroll_anim.target_x;
                self.ribbon_offset_target = scroll_anim.target_x;
                self.vertical_offset = scroll_anim.target_y;
                self.vertical_offset_target = scroll_anim.target_y;
                self.scroll_animation = None;
                self.focus_visible_window();
                self.needs_ribbon_recalc = true;
            } else {
                let t = elapsed.as_secs_f32() / scroll_anim.duration.as_secs_f32();
                let eased_t = Self::ease_out_cubic(t);
                
                self.ribbon_offset = Self::lerp(scroll_anim.start_x, scroll_anim.target_x, eased_t);
                self.vertical_offset = Self::lerp(scroll_anim.start_y, scroll_anim.target_y, eased_t);
            }
            need_reposition = true;
        }

        // Collect ribbon repositions if needed
        if need_reposition {
            for window in self.windows.values() {
                if window.animation.is_none() {
                    let rect = self.ribbon_to_screen(&window.position);
                    window_updates.push((window.hwnd, rect));
                }
            }
        }

        // Update individual window animations
        for (hwnd_val, window) in self.windows.iter_mut() {
            if let Some(anim) = &window.animation {
                let elapsed = now.duration_since(anim.start_time);
                
                if elapsed >= anim.duration {
                    match anim.animation_type {
                        AnimationType::Exit => {
                            animations_complete.push(*hwnd_val);
                        }
                        _ => {
                            animations_complete.push(*hwnd_val);
                            window_updates.push((window.hwnd, anim.target_rect));
                        }
                    }
                } else {
                    let t = elapsed.as_secs_f32() / anim.duration.as_secs_f32();
                    let eased_t = Self::ease_out_cubic(t);
                    
                    let current_rect = match anim.animation_type {
                        AnimationType::Move => {
                            RECT {
                                left: Self::lerp(anim.start_rect.left, anim.target_rect.left, eased_t),
                                top: Self::lerp(anim.start_rect.top, anim.target_rect.top, eased_t),
                                right: Self::lerp(anim.start_rect.right, anim.target_rect.right, eased_t),
                                bottom: Self::lerp(anim.start_rect.bottom, anim.target_rect.bottom, eased_t),
                            }
                        }
                        AnimationType::Entry => {
                            let center_x = (anim.target_rect.left + anim.target_rect.right) / 2;
                            let center_y = (anim.target_rect.top + anim.target_rect.bottom) / 2;
                            let target_width = anim.target_rect.right - anim.target_rect.left;
                            let target_height = anim.target_rect.bottom - anim.target_rect.top;
                            
                            let scale = 0.1 + 0.9 * eased_t;
                            let current_width = (target_width as f32 * scale) as i32;
                            let current_height = (target_height as f32 * scale) as i32;
                            
                            RECT {
                                left: center_x - current_width / 2,
                                top: center_y - current_height / 2,
                                right: center_x + current_width / 2,
                                bottom: center_y + current_height / 2,
                            }
                        }
                        AnimationType::Exit => {
                            RECT {
                                left: Self::lerp(anim.start_rect.left, anim.target_rect.left, eased_t),
                                top: Self::lerp(anim.start_rect.top, anim.target_rect.top, eased_t),
                                right: Self::lerp(anim.start_rect.right, anim.target_rect.right, eased_t),
                                bottom: Self::lerp(anim.start_rect.bottom, anim.target_rect.bottom, eased_t),
                            }
                        }
                    };
                    
                    window_updates.push((window.hwnd, current_rect));
                }
            }
        }

        // Apply all window updates in a single batch
        self.batch_set_window_positions(&window_updates);

        // Handle animation completion
        let mut windows_to_remove = Vec::new();
        
        for &hwnd_val in &animations_complete {
            if let Some(window) = self.windows.get(&hwnd_val) {
                if let Some(anim) = &window.animation {
                    if anim.animation_type == AnimationType::Exit {
                        windows_to_remove.push((hwnd_val, window.clone(), anim.target_rect));
                    }
                }
            }
        }
        
        // Now remove and restore windows
        for (hwnd_val, window_copy, target_rect) in windows_to_remove {
            self.windows.remove(&hwnd_val);
            
            unsafe {
                SetWindowLongW(window_copy.hwnd, GWL_STYLE, window_copy.original_style.0 as i32);
                SetWindowLongW(window_copy.hwnd, GWL_EXSTYLE, window_copy.original_ex_style.0 as i32);
                
                if (window_copy.original_ex_style & WS_EX_LAYERED).0 == 0 {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(window_copy.hwnd, GWL_EXSTYLE) as u32);
                    SetWindowLongW(window_copy.hwnd, GWL_EXSTYLE, 
                        (ex_style.0 & !WS_EX_LAYERED.0) as i32);
                }
                
                SetWindowPos(
                    window_copy.hwnd,
                    HWND_TOP,
                    target_rect.left,
                    target_rect.top,
                    target_rect.right - target_rect.left,
                    target_rect.bottom - target_rect.top,
                    SWP_NOZORDER | SWP_FRAMECHANGED,
                ).ok();
                
                ShowWindow(window_copy.hwnd, SW_RESTORE);
            }
            
            self.reflow_ribbon();
            self.needs_ribbon_recalc = true;
        }
        
        for &hwnd_val in &animations_complete {
            if let Some(window) = self.windows.get_mut(&hwnd_val) {
                window.animation = None;
            }
        }

        // Check if all animations are complete
        let all_complete = self.scroll_animation.is_none() &&
            self.windows.values().all(|w| w.animation.is_none());
        
        if all_complete {
            *self.animation_running.lock().unwrap() = false;
            *self.animation_stop_requested.lock().unwrap() = true;
            
            if self.needs_ribbon_recalc {
                let now = Instant::now();
                if now.duration_since(self.last_ribbon_recalc).as_millis() > 500 {
                    self.recalculate_ribbon();
                }
            }
        }
    }

    // Batch window position updates for better performance
    fn batch_set_window_positions(&self, updates: &[(HWND, RECT)]) {
        if updates.is_empty() {
            return;
        }

        unsafe {
            match BeginDeferWindowPos(updates.len() as i32) {
                Ok(hdwp) => {
                    let mut hdwp_current = hdwp;
                    
                    for (hwnd, rect) in updates {
                        let width = rect.right - rect.left;
                        let height = rect.bottom - rect.top;
                        
                        // Check if window is visible on screen
                        if width > 0 && height > 0 && 
                           rect.left < self.monitor_width * 2 && rect.right > -self.monitor_width &&
                           rect.top < self.monitor_height * 2 && rect.bottom > -self.monitor_height {
                            match DeferWindowPos(
                                hdwp_current,
                                *hwnd,
                                HWND::default(),
                                rect.left,
                                rect.top,
                                width,
                                height,
                                SWP_NOZORDER | SWP_NOACTIVATE,
                            ) {
                                Ok(new_hdwp) => {
                                    hdwp_current = new_hdwp;
                                }
                                Err(_) => {
                                    // Continue with the current handle
                                }
                            }
                        }
                    }
                    
                    EndDeferWindowPos(hdwp_current).ok();
                }
                Err(_) => {
                    // Fall back to individual updates
                    for (hwnd, rect) in updates {
                        let width = rect.right - rect.left;
                        let height = rect.bottom - rect.top;
                        
                        if width > 0 && height > 0 &&
                           rect.left > -20000 && rect.top > -20000 && 
                           rect.right < 20000 && rect.bottom < 20000 {
                            SetWindowPos(
                                *hwnd,
                                HWND_TOP,
                                rect.left,
                                rect.top,
                                width,
                                height,
                                SWP_NOZORDER | SWP_NOACTIVATE,
                            ).ok();
                        }
                    }
                }
            }
        }
    }

    fn set_window_rect(hwnd: HWND, rect: &RECT) {
        unsafe {
            let width = rect.right - rect.left;
            let height = rect.bottom - rect.top;
            
            if width <= 0 || height <= 0 {
                return;
            }
            
            // Allow windows to be positioned off-screen for smooth scrolling
            // but prevent extreme values that could cause issues
            if rect.left < -20000 || rect.top < -20000 || rect.right > 20000 || rect.bottom > 20000 {
                return;
            }
            
            SetWindowPos(
                hwnd,
                HWND_TOP,
                rect.left,
                rect.top,
                width,
                height,
                SWP_NOZORDER | SWP_NOACTIVATE,
            ).ok();
        }
    }

    // Start animation timer if not already running
    fn start_animation_timer(&self) {
        let mut running = self.animation_running.lock().unwrap();
        if !*running {
            *running = true;
            *self.animation_stop_requested.lock().unwrap() = false;
            
            let animation_running = self.animation_running.clone();
            let animation_stop_requested = self.animation_stop_requested.clone();
            let main_hwnd = self.main_hwnd;
            
            thread::spawn(move || {
                let start_time = Instant::now();
                let mut last_recalc_check = Instant::now();
                
                loop {
                    if *animation_stop_requested.lock().unwrap() {
                        break;
                    }
                    
                    if *animation_running.lock().unwrap() {
                        unsafe {
                            PostMessageW(
                                main_hwnd,
                                WM_USER + 1,
                                WPARAM(0),
                                LPARAM(0)
                            ).ok();
                            
                            let now = Instant::now();
                            if now.duration_since(last_recalc_check).as_millis() > 500 {
                                PostMessageW(
                                    main_hwnd,
                                    WM_TILER_RECALC,
                                    WPARAM(0),
                                    LPARAM(0)
                                ).ok();
                                last_recalc_check = now;
                            }
                        }
                        
                        let elapsed = Instant::now().duration_since(start_time);
                        let interval = if elapsed.as_millis() < 200 {
                            16  // 60 FPS
                        } else {
                            let extra = ((elapsed.as_millis() - 200) / 100).min(4);
                            16 + extra as u64
                        };
                        
                        thread::sleep(Duration::from_millis(interval));
                    } else {
                        break;
                    }
                }
                
                *animation_running.lock().unwrap() = false;
                *animation_stop_requested.lock().unwrap() = false;
            });
        }
    }

    fn ribbon_to_screen(&self, pos: &RibbonPosition) -> RECT {
        let base_x = pos.x - self.ribbon_offset;
        let base_y = pos.row * self.row_height - self.vertical_offset;
        
        let w = match pos.size {
            TileSize::Full => self.monitor_width,
            TileSize::Half => self.monitor_width / 2,
        };

        RECT {
            left: base_x + self.margin_horizontal / 2,
            top: base_y + self.margin_vertical / 2,
            right: base_x + w - self.margin_horizontal / 2,
            bottom: base_y + self.row_height - self.margin_vertical / 2,
        }
    }

    fn should_manage_window(&self, hwnd: HWND) -> bool {
        unsafe {
            if !IsWindowVisible(hwnd).as_bool() {
                return false;
            }

            let style = WINDOW_STYLE(GetWindowLongW(hwnd, GWL_STYLE) as u32);
            if (style & WS_MINIMIZE).0 != 0 {
                return false;
            }

            let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
            if (ex_style & WS_EX_TOOLWINDOW).0 != 0 {
                return false;
            }

            let mut class_name = [0u16; 256];
            let class_len = GetClassNameW(hwnd, &mut class_name);
            if class_len == 0 {
                return true;
            }
            let class_str = String::from_utf16_lossy(&class_name[..class_len as usize]);

            let system_classes = [
                "Shell_TrayWnd", "Shell_SecondaryTrayWnd", "TaskListThumbnailWnd",
                "MSTaskSwWClass", "ForegroundStaging", "Windows.UI.Core.CoreWindow",
                "Progman", "WorkerW", "DV2ControlHost", "Button", "Static",
                "#32770", "ToolbarWindow32", "tooltips_class32", "ComboLBox",
            ];

            if system_classes.iter().any(|&sc| class_str == sc) {
                return false;
            }

            let mut title = [0u16; 256];
            let len = GetWindowTextW(hwnd, &mut title);
            let title_str = String::from_utf16_lossy(&title[..len as usize]);

            let system_titles = ["Program Manager", "Task Switching", "Start"];
            if system_titles.iter().any(|&st| title_str.starts_with(st)) {
                return false;
            }

            if len == 0 && !class_str.contains("Chrome") && !class_str.contains("Firefox") {
                return false;
            }
            
            if (style & WS_VISIBLE).0 == 0 {
                return false;
            }

            if (style & WS_CAPTION).0 == 0 && (style & WS_POPUP).0 != 0 {
                return false;
            }

            if class_str == "ApplicationFrameWindow" && len > 0 {
                return true;
            }

            println!("Window added to ribbon (row {})", self.current_row);
            
            true
        }
    }

    fn is_popup_window(&self, hwnd: HWND) -> bool {
        unsafe {
            let style = WINDOW_STYLE(GetWindowLongW(hwnd, GWL_STYLE) as u32);
            let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
            
            let is_popup_no_caption = (style & WS_POPUP).0 != 0 && (style & WS_CAPTION).0 == 0;
            let is_dialog = (ex_style & WS_EX_DLGMODALFRAME).0 != 0;
            
            let owner = GetWindow(hwnd, GW_OWNER);
            let has_owner = owner != HWND::default();
            
            let mut class_name = [0u16; 256];
            let class_len = GetClassNameW(hwnd, &mut class_name);
            let class_str = String::from_utf16_lossy(&class_name[..class_len as usize]);
            let is_dialog_class = class_str == "#32770";
            
            is_dialog_class || (is_popup_no_caption && has_owner) || (is_dialog && has_owner)
        }
    }

    fn track_floating_window(&mut self, hwnd: HWND) {
        if !self.floating_windows.contains_key(&hwnd.0) {
            self.floating_windows.insert(hwnd.0, hwnd);
            
            if self.transparency < 255 {
                unsafe {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
                    SetWindowLongW(hwnd, GWL_EXSTYLE, 
                        (ex_style.0 | WS_EX_LAYERED.0) as i32);
                    SetLayeredWindowAttributes(hwnd, COLORREF(0), self.transparency, LWA_ALPHA).ok();
                }
            }
        }
    }

    fn clean_minimized_windows(&mut self) {
        let mut minimized = Vec::new();
        
        unsafe {
            for (hwnd_val, _) in self.windows.iter() {
                let hwnd = HWND(*hwnd_val);
                let style = WINDOW_STYLE(GetWindowLongW(hwnd, GWL_STYLE) as u32);
                if (style & WS_MINIMIZE).0 != 0 {
                    minimized.push(*hwnd_val);
                }
            }
        }
        
        if !minimized.is_empty() {
            for hwnd_val in &minimized {
                self.windows.remove(hwnd_val);
            }
            self.needs_ribbon_recalc = true;
        }
    }

    // Clean up windows that were closed externally
    fn clean_closed_windows(&mut self) {
        let mut closed_windows = Vec::new();
        
        unsafe {
            for (hwnd_val, _) in self.windows.iter() {
                let hwnd = HWND(*hwnd_val);
                if !IsWindow(hwnd).as_bool() || !IsWindowVisible(hwnd).as_bool() {
                    closed_windows.push(*hwnd_val);
                }
            }
        }
        
        if !closed_windows.is_empty() {
            for hwnd_val in &closed_windows {
                self.windows.remove(hwnd_val);
            }
            self.needs_ribbon_recalc = true;
        }
    }

    // Get all rows that have windows
    fn get_active_rows(&self) -> Vec<i32> {
        let mut rows: Vec<i32> = self.windows.values()
            .map(|w| w.position.row)
            .collect();
        rows.sort();
        rows.dedup();
        rows
    }

    // Check if any row has a window at the given x position
    fn is_x_position_occupied(&self, x: i32, width: i32) -> bool {
        self.windows.values().any(|w| {
            let window_start = w.position.x;
            let window_end = w.position.x + self.get_tile_width(&w.position.size);
            // Check if ranges overlap
            !(x + width <= window_start || x >= window_end)
        })
    }

    // Recalculate entire ribbon layout
    fn recalculate_ribbon(&mut self) {
        self.clean_closed_windows();
        
        if self.windows.is_empty() {
            self.ribbon_offset = 0;
            self.ribbon_offset_target = 0;
            self.vertical_offset = 0;
            self.vertical_offset_target = 0;
            return;
        }
        
        // Get all non-exiting windows grouped by row
        let mut rows: HashMap<i32, Vec<(isize, RibbonPosition)>> = HashMap::new();
        
        for (hwnd, w) in self.windows.iter() {
            if w.animation.as_ref().map_or(true, |a| a.animation_type != AnimationType::Exit) {
                rows.entry(w.position.row)
                    .or_insert_with(Vec::new)
                    .push((*hwnd, w.position));
            }
        }
        
        // Sort each row by x position
        for windows in rows.values_mut() {
            windows.sort_by_key(|(_, pos)| pos.x);
        }
        
        // Find all unique x positions across all rows
        let mut x_positions: Vec<(i32, i32)> = Vec::new(); // (x, width)
        
        for windows in rows.values() {
            for (_, pos) in windows {
                x_positions.push((pos.x, self.get_tile_width(&pos.size)));
            }
        }
        
        // Sort and deduplicate x positions
        x_positions.sort_by_key(|(x, _)| *x);
        
        // Build a new layout with no gaps
        let mut new_x_mapping: HashMap<i32, i32> = HashMap::new();
        let mut current_x = 0;
        
        for (old_x, width) in x_positions {
            if !new_x_mapping.contains_key(&old_x) {
                new_x_mapping.insert(old_x, current_x);
                // Only advance if this x position is actually used
                if self.is_x_position_occupied(old_x, width) {
                    current_x += width;
                }
            }
        }
        
        // Update all window positions
        let mut positions_to_update = Vec::new();
        
        for (_hwnd, window) in self.windows.iter_mut() {
            if let Some(&new_x) = new_x_mapping.get(&window.position.x) {
                window.position.x = new_x;
                
                if window.animation.is_none() && window.position.row == self.current_row {
                    positions_to_update.push((window.hwnd, window.position));
                }
            }
        }
        
        // Apply position updates
        for (hwnd, position) in positions_to_update {
            let rect = self.ribbon_to_screen(&position);
            Self::set_window_rect(hwnd, &rect);
        }
        
        // Ensure ribbon offset is within bounds
        let max_x = self.windows.values()
            .map(|w| w.position.x + self.get_tile_width(&w.position.size))
            .max()
            .unwrap_or(0);
        let max_offset = (max_x - self.monitor_width).max(0);
        
        if self.ribbon_offset > max_offset {
            self.ribbon_offset = max_offset;
            self.ribbon_offset_target = max_offset;
            self.scroll_animation = None;
        }
        
        self.needs_ribbon_recalc = false;
        self.last_ribbon_recalc = Instant::now();
    }

    fn reflow_ribbon(&mut self) {
        self.needs_ribbon_recalc = true;
    }

    fn add_window(&mut self, hwnd: HWND) -> bool {
        self.check_monitor_dimensions();
        
        if self.windows.contains_key(&hwnd.0) {
            return false;
        }
        
        if !self.should_manage_window(hwnd) {
            return false;
        }

        if self.is_popup_window(hwnd) {
            self.track_floating_window(hwnd);
            unsafe {
                SetForegroundWindow(hwnd);
            }
            return true;
        }

        unsafe {
            let style = WINDOW_STYLE(GetWindowLongW(hwnd, GWL_STYLE) as u32);
            let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
            
            let mut rect = RECT::default();
            GetWindowRect(hwnd, &mut rect).ok();
            
            let screen_width = GetSystemMetrics(SM_CXSCREEN);
            let screen_height = GetSystemMetrics(SM_CYSCREEN);
            let width = rect.right - rect.left;
            let height = rect.bottom - rect.top;
            
            if rect.left < -width + 100 || rect.left > screen_width - 100 ||
               rect.top < -height + 100 || rect.top > screen_height - 100 {
                rect.left = (screen_width - width) / 2;
                rect.top = (screen_height - height) / 2;
                rect.right = rect.left + width;
                rect.bottom = rect.top + height;
            }
            
            if IsZoomed(hwnd).as_bool() {
                ShowWindow(hwnd, SW_RESTORE);
            }
            
            let new_style = WINDOW_STYLE(style.0 & !WS_MINIMIZEBOX.0 & !WS_MAXIMIZEBOX.0 & !WS_MAXIMIZE.0);
            SetWindowLongW(hwnd, GWL_STYLE, new_style.0 as i32);
            
            SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, 
                SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOACTIVATE);
            
            let position = self.find_viewport_position();
            
            let window = ManagedWindow {
                hwnd,
                original_style: style,
                original_ex_style: ex_style,
                original_rect: rect,
                position,
                animation: None,
            };

            self.windows.insert(hwnd.0, window);
            
            let new_window_width = self.get_tile_width(&position.size);
            let insertion_x = position.x;
            
            // Shift windows on the same row
            let windows_to_shift: Vec<isize> = self.windows.iter()
                .filter(|(h, w)| **h != hwnd.0 && w.position.row == position.row && w.position.x >= insertion_x)
                .map(|(h, _)| *h)
                .collect();
                
            for hwnd_to_shift in windows_to_shift {
                if let Some(w) = self.windows.get_mut(&hwnd_to_shift) {
                    w.position.x += new_window_width;
                }
            }
            
            let window_end = insertion_x + new_window_width;
            if insertion_x < self.ribbon_offset || window_end > self.ribbon_offset + self.monitor_width {
                let center_offset = insertion_x + new_window_width / 2 - self.monitor_width / 2;
                
                let max_x = self.windows.values()
                    .map(|w| w.position.x + self.get_tile_width(&w.position.size))
                    .max()
                    .unwrap_or(0);
                let max_offset = (max_x - self.monitor_width).max(0);
                
                self.ribbon_offset = center_offset.clamp(0, max_offset);
                self.ribbon_offset_target = self.ribbon_offset;
                self.vertical_offset = self.current_row * self.row_height;
                self.vertical_offset_target = self.vertical_offset;
            }
            
            self.apply_window_position_with_animation_type(hwnd, AnimationType::Entry);
            
            let shifted_hwnds: Vec<HWND> = self.windows.iter()
                .filter(|(h, w)| **h != hwnd.0 && w.position.row == position.row && w.position.x >= insertion_x + new_window_width)
                .map(|(_, w)| w.hwnd)
                .collect();
            
            for shifted_hwnd in shifted_hwnds {
                self.apply_window_position(shifted_hwnd, true);
            }
            
            self.needs_ribbon_recalc = true;
            
            true
        }
    }

    fn find_viewport_position(&self) -> RibbonPosition {
        let focused_hwnd = unsafe { GetForegroundWindow() };
        let focused_center = self.windows.get(&focused_hwnd.0)
            .map(|w| w.position.x + self.get_tile_width(&w.position.size) / 2)
            .unwrap_or_else(|| self.ribbon_offset + self.monitor_width / 2);
        
        let mut best_position = self.ribbon_offset;
        let mut best_distance = i32::MAX;
        
        // Only check windows on the current row
        for window in self.windows.values().filter(|w| w.position.row == self.current_row) {
            let left_edge = window.position.x;
            let right_edge = window.position.x + self.get_tile_width(&window.position.size);
            
            let left_distance = (left_edge - focused_center).abs();
            if left_distance < best_distance {
                best_distance = left_distance;
                best_position = left_edge;
            }
            
            let right_distance = (right_edge - focused_center).abs();
            if right_distance < best_distance {
                best_distance = right_distance;
                best_position = right_edge;
            }
        }
        
        RibbonPosition {
            x: best_position,
            row: self.current_row,
            size: TileSize::Half,
        }
    }

    fn remove_window(&mut self, hwnd: HWND) {
        if let Some(window) = self.windows.get_mut(&hwnd.0) {
            let mut current_rect = RECT::default();
            unsafe {
                GetWindowRect(hwnd, &mut current_rect).ok();
            }
            
            let width = window.original_rect.right - window.original_rect.left;
            let height = window.original_rect.bottom - window.original_rect.top;
            
            let screen_width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
            let screen_height = unsafe { GetSystemMetrics(SM_CYSCREEN) };
            
            let mut left = window.original_rect.left;
            let mut top = window.original_rect.top;
            
            if left < -width + 100 || 
               left > screen_width - 100 ||
               top < -height + 100 || 
               top > screen_height - 100 {
                left = (screen_width - width) / 2;
                top = (screen_height - height) / 2;
            }
            
            left = left.clamp(-width + 100, screen_width - 100);
            top = top.clamp(-height + 100, screen_height - 100);
            
            let target_rect = RECT {
                left,
                top,
                right: left + width,
                bottom: top + height,
            };
            
            window.animation = Some(AnimationState {
                start_rect: current_rect,
                target_rect,
                start_time: Instant::now(),
                duration: Duration::from_millis(200),
                animation_type: AnimationType::Exit,
            });
            
            self.start_animation_timer();
        }
    }

    fn restore_window(&self, window: &ManagedWindow) {
        unsafe {
            SetWindowLongW(window.hwnd, GWL_STYLE, window.original_style.0 as i32);
            SetWindowLongW(window.hwnd, GWL_EXSTYLE, window.original_ex_style.0 as i32);
            
            if (window.original_ex_style & WS_EX_LAYERED).0 == 0 {
                let ex_style = WINDOW_EX_STYLE(GetWindowLongW(window.hwnd, GWL_EXSTYLE) as u32);
                SetWindowLongW(window.hwnd, GWL_EXSTYLE, 
                    (ex_style.0 & !WS_EX_LAYERED.0) as i32);
            }
            
            let width = window.original_rect.right - window.original_rect.left;
            let height = window.original_rect.bottom - window.original_rect.top;
            
            let screen_width = GetSystemMetrics(SM_CXSCREEN);
            let screen_height = GetSystemMetrics(SM_CYSCREEN);
            
            let mut left = window.original_rect.left;
            let mut top = window.original_rect.top;
            
            if left < -width + 100 || 
               left > screen_width - 100 ||
               top < -height + 100 || 
               top > screen_height - 100 {
                left = (screen_width - width) / 2;
                top = (screen_height - height) / 2;
            }
            
            left = left.clamp(-width + 100, screen_width - 100);
            top = top.clamp(-height + 100, screen_height - 100);
            
            SetWindowPos(
                window.hwnd,
                HWND_TOP,
                left,
                top,
                width,
                height,
                SWP_NOZORDER | SWP_FRAMECHANGED,
            ).ok();
            
            ShowWindow(window.hwnd, SW_RESTORE);
        }
    }

    fn shutdown(&mut self) {
        println!("\nShutting down Thymeline...");
        
        *self.animation_stop_requested.lock().unwrap() = true;
        
        let screen_width = unsafe { GetSystemMetrics(SM_CXSCREEN) };
        let screen_height = unsafe { GetSystemMetrics(SM_CYSCREEN) };
        
        for window in self.windows.values_mut() {
            let mut current_rect = RECT::default();
            unsafe {
                GetWindowRect(window.hwnd, &mut current_rect).ok();
            }
            
            let width = window.original_rect.right - window.original_rect.left;
            let height = window.original_rect.bottom - window.original_rect.top;
            
            let mut left = window.original_rect.left;
            let mut top = window.original_rect.top;
            
            if left < -width + 100 || 
               left > screen_width - 100 ||
               top < -height + 100 || 
               top > screen_height - 100 {
                left = (screen_width - width) / 2;
                top = (screen_height - height) / 2;
            }
            
            left = left.clamp(-width + 100, screen_width - 100);
            top = top.clamp(-height + 100, screen_height - 100);
            
            let target_rect = RECT {
                left,
                top,
                right: left + width,
                bottom: top + height,
            };
            
            window.animation = Some(AnimationState {
                start_rect: current_rect,
                target_rect,
                start_time: Instant::now(),
                duration: Duration::from_millis(150),
                animation_type: AnimationType::Exit,
            });
        }
        
        self.start_animation_timer();
        
        thread::sleep(Duration::from_millis(200));
        
        let windows: Vec<ManagedWindow> = self.windows.values().cloned().collect();
        for window in windows {
            self.restore_window(&window);
        }
        
        self.windows.clear();
        self.ribbon_offset = 0;
        self.ribbon_offset_target = 0;
        self.vertical_offset = 0;
        self.vertical_offset_target = 0;
        self.current_row = 0;
        self.scroll_animation = None;
        
        for (_, hwnd) in &self.floating_windows {
            unsafe {
                if IsWindow(*hwnd).as_bool() {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(*hwnd, GWL_EXSTYLE) as u32);
                    if self.transparency < 255 {
                        SetWindowLongW(*hwnd, GWL_EXSTYLE, 
                            (ex_style.0 & !WS_EX_LAYERED.0) as i32);
                        SetWindowPos(*hwnd, HWND_TOP, 0, 0, 0, 0,
                            SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOZORDER).ok();
                    }
                }
            }
        }
        
        self.floating_windows.clear();
        
        println!("All windows restored to original state");
    }

    fn apply_window_position(&mut self, hwnd: HWND, animate: bool) {
        if animate {
            self.apply_window_position_with_animation_type(hwnd, AnimationType::Move);
        } else {
            let position = match self.windows.get(&hwnd.0) {
                Some(window) => window.position,
                None => return,
            };
            
            let target_rect = self.ribbon_to_screen(&position);
            
            unsafe {
                if let Some(window) = self.windows.get(&hwnd.0) {
                    if IsZoomed(hwnd).as_bool() {
                        ShowWindow(hwnd, SW_RESTORE);
                    }
                    
                    let style = window.original_style;
                    let style = WINDOW_STYLE(style.0 & !WS_MINIMIZEBOX.0 & !WS_MAXIMIZEBOX.0 & !WS_MAXIMIZE.0);
                    SetWindowLongW(hwnd, GWL_STYLE, style.0 as i32);
                    SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, 
                        SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOACTIVATE);
                }

                if self.transparency < 255 {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
                    SetWindowLongW(hwnd, GWL_EXSTYLE, 
                        (ex_style.0 | WS_EX_LAYERED.0) as i32);
                    SetLayeredWindowAttributes(hwnd, COLORREF(0), self.transparency, LWA_ALPHA).ok();
                } else {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
                    SetWindowLongW(hwnd, GWL_EXSTYLE, 
                        (ex_style.0 & !WS_EX_LAYERED.0) as i32);
                }
                
                ShowWindow(hwnd, SW_RESTORE);
                
                SetWindowPos(
                    hwnd,
                    HWND_TOP,
                    target_rect.left,
                    target_rect.top,
                    target_rect.right - target_rect.left,
                    target_rect.bottom - target_rect.top,
                    SWP_NOZORDER | SWP_NOACTIVATE,
                ).ok();
            }
        }
    }
    
    fn apply_window_position_with_animation_type(&mut self, hwnd: HWND, animation_type: AnimationType) {
        let position = match self.windows.get(&hwnd.0) {
            Some(window) => window.position,
            None => return,
        };
        
        let target_rect = self.ribbon_to_screen(&position);
        
        unsafe {
            if let Some(window) = self.windows.get(&hwnd.0) {
                if IsZoomed(hwnd).as_bool() {
                    ShowWindow(hwnd, SW_RESTORE);
                }
                
                let style = window.original_style;
                let style = WINDOW_STYLE(style.0 & !WS_MINIMIZEBOX.0 & !WS_MAXIMIZEBOX.0 & !WS_MAXIMIZE.0);
                SetWindowLongW(hwnd, GWL_STYLE, style.0 as i32);
                SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0, 
                    SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED | SWP_NOACTIVATE);
            }

            if self.transparency < 255 {
                let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
                SetWindowLongW(hwnd, GWL_EXSTYLE, 
                    (ex_style.0 | WS_EX_LAYERED.0) as i32);
                SetLayeredWindowAttributes(hwnd, COLORREF(0), self.transparency, LWA_ALPHA).ok();
            } else {
                let ex_style = WINDOW_EX_STYLE(GetWindowLongW(hwnd, GWL_EXSTYLE) as u32);
                SetWindowLongW(hwnd, GWL_EXSTYLE, 
                    (ex_style.0 & !WS_EX_LAYERED.0) as i32);
            }
        }
        
        if let Some(window) = self.windows.get_mut(&hwnd.0) {
            unsafe {
                let mut current_rect = RECT::default();
                GetWindowRect(hwnd, &mut current_rect).ok();
                
                if animation_type == AnimationType::Entry {
                    current_rect = target_rect;
                    ShowWindow(hwnd, SW_RESTORE);
                } else if position.row == self.current_row {
                    ShowWindow(hwnd, SW_RESTORE);
                }
                
                let duration = match animation_type {
                    AnimationType::Entry => Duration::from_millis(200),
                    AnimationType::Exit => Duration::from_millis(200),
                    AnimationType::Move => Duration::from_millis(87),
                };
                
                window.animation = Some(AnimationState {
                    start_rect: current_rect,
                    target_rect,
                    start_time: Instant::now(),
                    duration,
                    animation_type,
                });
                
                drop(window);
                self.start_animation_timer();
            }
        }
    }

    fn find_next_free_position(&self) -> RibbonPosition {
        let max_x = self.windows.values()
            .filter(|w| w.position.row == self.current_row)
            .map(|w| w.position.x + self.get_tile_width(&w.position.size))
            .max()
            .unwrap_or(0);

        let x = if self.windows.is_empty() {
            self.ribbon_offset
        } else {
            max_x
        };

        RibbonPosition {
            x,
            row: self.current_row,
            size: TileSize::Half,
        }
    }

    fn get_tile_width(&self, size: &TileSize) -> i32 {
        match size {
            TileSize::Full => self.monitor_width,
            TileSize::Half => self.monitor_width / 2,
        }
    }

    fn resize_window(&mut self, hwnd: HWND, direction: Direction) {
        self.check_monitor_dimensions();
        self.clean_closed_windows();
        self.clean_minimized_windows();
        
        if !self.windows.contains_key(&hwnd.0) {
            if !self.add_window(hwnd) {
                return;
            }
        }

        if let Some(window) = self.windows.get(&hwnd.0).cloned() {
            let old_size = window.position.size;
            let old_width = self.get_tile_width(&old_size);
            
            let new_size = match (old_size, direction) {
                (TileSize::Full, Direction::Left | Direction::Right) => TileSize::Half,
                (TileSize::Half, Direction::Left | Direction::Right) => TileSize::Full,
                _ => old_size,
            };
            
            let new_width = self.get_tile_width(&new_size);
            let width_diff = new_width - old_width;
            
            if let Some(w) = self.windows.get_mut(&hwnd.0) {
                w.position.size = new_size;
            }
            
            // If expanding, push windows to the right
            if width_diff > 0 {
                let current_pos = window.position;
                let current_end = current_pos.x + old_width;
                
                // Find all windows to the right on the same row that need to be pushed
                let windows_to_push: Vec<isize> = self.windows.iter()
                    .filter(|(h, w)| {
                        **h != hwnd.0 && 
                        w.position.row == current_pos.row && 
                        w.position.x >= current_end
                    })
                    .map(|(h, _)| *h)
                    .collect();
                
                // Push them right by the width difference
                for hwnd_to_push in windows_to_push {
                    if let Some(w) = self.windows.get_mut(&hwnd_to_push) {
                        w.position.x += width_diff;
                    }
                }
            }
            
            self.needs_ribbon_recalc = true;
        }
        
        self.apply_all_windows(true);
    }

    fn pull_adjacent_windows(&mut self, _changed_hwnd: isize) {
        self.needs_ribbon_recalc = true;
    }

    // Move window between rows or swap positions
    fn move_window(&mut self, hwnd: HWND, direction: Direction) {
        self.check_monitor_dimensions();
        self.clean_closed_windows();
        self.clean_minimized_windows();
        
        if !self.windows.contains_key(&hwnd.0) {
            return;
        }

        // Clear any existing animation on the focused window
        if let Some(window) = self.windows.get_mut(&hwnd.0) {
            window.animation = None;
        }

        let current_pos = match self.windows.get(&hwnd.0) {
            Some(w) => w.position,
            None => return,
        };
        
        match direction {
            Direction::Up | Direction::Down => {
                let old_row = current_pos.row;
                let new_row = match direction {
                    Direction::Up => {
                        if current_pos.row > 0 {
                            current_pos.row - 1
                        } else {
                            return;
                        }
                    },
                    Direction::Down => current_pos.row + 1,
                    _ => unreachable!(),
                };
                
                let current_width = self.get_tile_width(&current_pos.size);
                let current_x = current_pos.x;
                let current_end = current_x + current_width;
                
                // Check if the target position is empty
                let mut is_empty = true;
                let mut blocking_windows = Vec::new();
                
                for (other_hwnd, window) in self.windows.iter() {
                    if window.position.row == new_row {
                        let other_x = window.position.x;
                        let other_width = self.get_tile_width(&window.position.size);
                        let other_end = other_x + other_width;
                        
                        // Check if this window overlaps with our target position
                        if !(current_end <= other_x || current_x >= other_end) {
                            is_empty = false;
                            blocking_windows.push((*other_hwnd, other_x, other_width));
                        }
                    }
                }
                
                // Store old positions for animation
                let old_ribbon_offset = self.ribbon_offset;
                let old_vertical_offset = self.vertical_offset;
                
                if is_empty {
                    // Just move to the empty space
                    if let Some(window) = self.windows.get_mut(&hwnd.0) {
                        window.position.row = new_row;
                    }
                } else {
                    // Need to swap or shift windows
                    blocking_windows.sort_by_key(|(_, x, _)| *x);
                    
                    // Shift blocking windows to make room
                    for (other_hwnd, _, _) in &blocking_windows {
                        if let Some(w) = self.windows.get_mut(other_hwnd) {
                            w.position.x += current_width;
                        }
                    }
                    
                    // Move the current window to the new row
                    if let Some(window) = self.windows.get_mut(&hwnd.0) {
                        window.position.row = new_row;
                    }
                }
                
                // Update viewport to keep focused window stationary
                let row_diff = new_row - old_row;
                self.current_row = new_row;
                self.vertical_offset = old_vertical_offset + row_diff * self.row_height;
                self.vertical_offset_target = self.vertical_offset;
                
                // Start smooth universe movement
                self.animate_universe_movement(hwnd, old_ribbon_offset, old_vertical_offset);
                
                self.needs_ribbon_recalc = true;
                
                unsafe {
                    SetForegroundWindow(hwnd);
                }
            },
            Direction::Left | Direction::Right => {
                // Define the step size - half monitor width for consistent increments
                let step_size = self.monitor_width / 2;
                
                let current_width = self.get_tile_width(&current_pos.size);
                let old_x = current_pos.x;
                let mut new_x = old_x;
                
                // Get windows on the same row for collision checking
                let windows_on_row: Vec<(isize, i32, i32)> = self.windows.iter()
                    .filter(|(_, w)| w.position.row == current_pos.row)
                    .map(|(h, w)| (*h, w.position.x, self.get_tile_width(&w.position.size)))
                    .collect();
                
                match direction {
                    Direction::Left => {
                        // Move left by step_size, but check for collisions
                        let proposed_x = (old_x - step_size).max(0);
                        
                        // Check if this position would overlap with any window
                        let mut blocked = false;
                        for (other_hwnd, other_x, other_width) in &windows_on_row {
                            if *other_hwnd != hwnd.0 {
                                let other_end = other_x + other_width;
                                // Check if we would overlap
                                if proposed_x < other_end && proposed_x + current_width > *other_x {
                                    // We would overlap, so swap positions instead
                                    new_x = *other_x;
                                    if let Some(w) = self.windows.get_mut(other_hwnd) {
                                        w.position.x = old_x;
                                    }
                                    blocked = true;
                                    break;
                                }
                            }
                        }
                        
                        if !blocked {
                            new_x = proposed_x;
                        }
                    },
                    Direction::Right => {
                        // Move right by step_size
                        let proposed_x = old_x + step_size;
                        
                        // Check if this position would overlap with any window
                        let mut blocked = false;
                        for (other_hwnd, other_x, other_width) in &windows_on_row {
                            if *other_hwnd != hwnd.0 {
                                let other_end = other_x + other_width;
                                // Check if we would overlap
                                if proposed_x < other_end && proposed_x + current_width > *other_x {
                                    // We would overlap, so swap positions instead
                                    new_x = *other_x;
                                    if let Some(w) = self.windows.get_mut(other_hwnd) {
                                        w.position.x = old_x;
                                    }
                                    blocked = true;
                                    break;
                                }
                            }
                        }
                        
                        if !blocked {
                            new_x = proposed_x;
                        }
                    },
                    _ => unreachable!(),
                }
                
                // Calculate actual movement distance
                let movement_distance = new_x - old_x;
                
                if movement_distance == 0 {
                    return; // No movement needed
                }
                
                // Store old offset for animation
                let old_ribbon_offset = self.ribbon_offset;
                let old_vertical_offset = self.vertical_offset;
                
                // Update position
                if let Some(w) = self.windows.get_mut(&hwnd.0) {
                    w.position.x = new_x;
                }
                
                // Update ribbon offset to keep focused window stationary
                self.ribbon_offset = old_ribbon_offset + movement_distance;
                self.ribbon_offset_target = self.ribbon_offset;
                
                // Clamp to valid bounds
                let max_x = self.windows.values()
                    .map(|w| w.position.x + self.get_tile_width(&w.position.size))
                    .max()
                    .unwrap_or(0);
                let max_offset = (max_x - self.monitor_width).max(0);
                self.ribbon_offset = self.ribbon_offset.clamp(0, max_offset);
                self.ribbon_offset_target = self.ribbon_offset;
                
                // Start smooth universe movement
                self.animate_universe_movement(hwnd, old_ribbon_offset, old_vertical_offset);
                
                self.needs_ribbon_recalc = true;
                
                unsafe {
                    SetForegroundWindow(hwnd);
                }
            }
        }
    }
    
    // Smoothly animate universe movement from old viewport to new viewport
    fn animate_universe_movement(&mut self, focused_hwnd: HWND, old_ribbon_offset: i32, old_vertical_offset: i32) {
        // Extract values we'll need in the loop
        let new_ribbon_offset = self.ribbon_offset;
        let new_vertical_offset = self.vertical_offset;
        let margin_h = self.margin_horizontal;
        let margin_v = self.margin_vertical;
        let row_height = self.row_height;
        let monitor_width = self.monitor_width;
        
        // For each window, calculate where it would be with the OLD viewport
        // and where it should be with the NEW viewport, then animate between them
        for (hwnd_val, window) in self.windows.iter_mut() {
            if *hwnd_val == focused_hwnd.0 {
                // The focused window doesn't move
                continue;
            }
            
            // Get tile width before we need it
            let tile_width = match window.position.size {
                TileSize::Full => monitor_width,
                TileSize::Half => monitor_width / 2,
            };
            
            // Calculate old screen position (with old viewport)
            let old_screen_x = window.position.x - old_ribbon_offset;
            let old_screen_y = window.position.row * row_height - old_vertical_offset;
            
            // Calculate new screen position (with new viewport)
            let new_screen_x = window.position.x - new_ribbon_offset;
            let new_screen_y = window.position.row * row_height - new_vertical_offset;
            
            // If there's already an animation in progress, we need to handle it carefully
            let start_rect = if let Some(existing_anim) = &window.animation {
                // Use the current interpolated position as the start
                let elapsed = Instant::now().duration_since(existing_anim.start_time);
                let t = (elapsed.as_secs_f32() / existing_anim.duration.as_secs_f32()).min(1.0);
                let eased_t = Self::ease_out_cubic(t);
                
                RECT {
                    left: Self::lerp(existing_anim.start_rect.left, existing_anim.target_rect.left, eased_t),
                    top: Self::lerp(existing_anim.start_rect.top, existing_anim.target_rect.top, eased_t),
                    right: Self::lerp(existing_anim.start_rect.right, existing_anim.target_rect.right, eased_t),
                    bottom: Self::lerp(existing_anim.start_rect.bottom, existing_anim.target_rect.bottom, eased_t),
                }
            } else {
                // Use the old screen position
                RECT {
                    left: old_screen_x + margin_h / 2,
                    top: old_screen_y + margin_v / 2,
                    right: old_screen_x + tile_width - margin_h / 2,
                    bottom: old_screen_y + row_height - margin_v / 2,
                }
            };
            
            // Target is the new screen position
            let target_rect = RECT {
                left: new_screen_x + margin_h / 2,
                top: new_screen_y + margin_v / 2,
                right: new_screen_x + tile_width - margin_h / 2,
                bottom: new_screen_y + row_height - margin_v / 2,
            };
            
            // Create smooth animation
            window.animation = Some(AnimationState {
                start_rect,
                target_rect,
                start_time: Instant::now(),
                duration: Duration::from_millis(200),
                animation_type: AnimationType::Move,
            });
        }
        
        self.start_animation_timer();
    }

    fn recalculate_positions_for_new_resolution(&mut self, old_width: i32) {
        if self.windows.is_empty() {
            return;
        }
        
        let scale_factor = self.monitor_width as f32 / old_width as f32;
        
        // Recalculate row height
        self.row_height = self.monitor_height;
        
        // Group windows by row
        let mut rows: HashMap<i32, Vec<(isize, RibbonPosition)>> = HashMap::new();
        for (hwnd, w) in self.windows.iter() {
            rows.entry(w.position.row)
                .or_insert_with(Vec::new)
                .push((*hwnd, w.position));
        }
        
        // Process each row
        for (row, mut windows) in rows {
            windows.sort_by_key(|(_, pos)| pos.x);
            
            let mut current_x = 0;
            for (hwnd, old_pos) in windows {
                if let Some(window) = self.windows.get_mut(&hwnd) {
                    window.position.x = current_x;
                    window.position.row = row;
                    window.position.size = old_pos.size;
                    
                    current_x += self.get_tile_width(&old_pos.size);
                }
            }
        }
        
        self.ribbon_offset = (self.ribbon_offset as f32 * scale_factor) as i32;
        self.ribbon_offset_target = self.ribbon_offset;
        self.vertical_offset = self.current_row * self.row_height;
        self.vertical_offset_target = self.vertical_offset;
        
        self.needs_ribbon_recalc = true;
    }

    fn apply_all_windows(&mut self, animate: bool) {
        let hwnds: Vec<HWND> = self.windows.values()
            .map(|w| w.hwnd)
            .collect();
        
        for hwnd in hwnds {
            self.apply_window_position(hwnd, animate);
        }
    }

    fn check_monitor_dimensions(&mut self) {
        let now = Instant::now();
        if now.duration_since(self.last_resolution_check).as_millis() < self.resolution_check_throttle_ms as u128 {
            return;
        }
        
        self.last_resolution_check = now;
        
        let (new_width, new_height) = Self::get_monitor_dimensions();
        
        if new_width != self.monitor_width || new_height != self.monitor_height {
            let old_width = self.monitor_width;
            self.monitor_width = new_width;
            self.monitor_height = new_height;
            
            self.recalculate_positions_for_new_resolution(old_width);
            
            self.apply_all_windows(false);
        }
    }

    fn pan_ribbon(&mut self, direction: Direction) {
        self.check_monitor_dimensions();
        self.clean_closed_windows();
        
        if self.windows.is_empty() {
            return;
        }
        
        let max_x = self.windows.values()
            .map(|w| w.position.x + self.get_tile_width(&w.position.size))
            .max()
            .unwrap_or(0);
        let max_offset = (max_x - self.monitor_width).max(0);
        
        // Check if we're already at the edge
        match direction {
            Direction::Left if self.ribbon_offset_target <= 0 => return,
            Direction::Right if self.ribbon_offset_target >= max_offset => return,
            _ => {},
        }
        
        let snap_distance = self.monitor_width / 2;
        
        match direction {
            Direction::Left => {
                self.ribbon_offset_target = (self.ribbon_offset_target - snap_distance).max(0);
            },
            Direction::Right => {
                self.ribbon_offset_target = (self.ribbon_offset_target + snap_distance).min(max_offset);
            },
            _ => return,
        }
        
        self.start_scroll_animation();
    }
    
    // Pan between rows
    fn pan_row(&mut self, direction: Direction) {
        self.check_monitor_dimensions();
        self.clean_closed_windows();
        
        // Get the maximum row that has windows
        let max_row_with_windows = self.windows.values()
            .map(|w| w.position.row)
            .max()
            .unwrap_or(0);
        
        // Allow panning one row beyond the last window row (for empty space)
        // but no further
        let max_allowed_row = max_row_with_windows + 1;
        
        match direction {
            Direction::Up => {
                if self.current_row > 0 {
                    self.current_row -= 1;
                    self.vertical_offset_target = self.current_row * self.row_height;
                    println!("Targeting row {}", self.current_row);
                    self.start_scroll_animation();
                }
            },
            Direction::Down => {
                if self.current_row < max_allowed_row {
                    self.current_row += 1;
                    self.vertical_offset_target = self.current_row * self.row_height;
                    println!("Targeting row {}", self.current_row);
                    self.start_scroll_animation();
                }
            },
            _ => return,
        };
    }
    
    // Start or update scroll animation to current targets
    fn start_scroll_animation(&mut self) {
        // If we're already animating, just update the targets
        // The animation will smoothly interpolate to the new destination
        
        // Clamp targets to valid bounds
        let max_row = self.windows.values()
            .map(|w| w.position.row)
            .max()
            .unwrap_or(0);
        let max_vertical = max_row * self.row_height;
        self.vertical_offset_target = self.vertical_offset_target.clamp(0, max_vertical);
        
        let max_x = self.windows.values()
            .map(|w| w.position.x + self.get_tile_width(&w.position.size))
            .max()
            .unwrap_or(0);
        let max_horizontal = (max_x - self.monitor_width).max(0);
        self.ribbon_offset_target = self.ribbon_offset_target.clamp(0, max_horizontal);
        
        // Start new animation from current position to target
        self.scroll_animation = Some(ScrollAnimation {
            start_x: self.ribbon_offset,
            start_y: self.vertical_offset,
            target_x: self.ribbon_offset_target,
            target_y: self.vertical_offset_target,
            start_time: Instant::now(),
            duration: Duration::from_millis(200), // Smooth animation duration
        });
        
        self.start_animation_timer();
    }
    
    fn focus_visible_window(&self) {
        unsafe {
            let mut best_window: Option<HWND> = None;
            let mut best_distance = f32::MAX;
            
            let screen_center_x = self.monitor_width as f32 / 2.0;
            let screen_center_y = self.monitor_height as f32 / 2.0;
            
            for window in self.windows.values() {
                if !self.is_window_visible(&window.position) {
                    continue;
                }
                
                let rect = self.ribbon_to_screen(&window.position);
                let window_center_x = ((rect.left + rect.right) / 2) as f32;
                let window_center_y = ((rect.top + rect.bottom) / 2) as f32;
                
                // Calculate distance from screen center
                let dx = window_center_x - screen_center_x;
                let dy = window_center_y - screen_center_y;
                let distance = (dx * dx + dy * dy).sqrt();
                
                if distance < best_distance {
                    best_distance = distance;
                    best_window = Some(window.hwnd);
                }
            }
            
            if let Some(hwnd) = best_window {
                let current_foreground = GetForegroundWindow();
                if current_foreground != hwnd {
                    SetActiveWindow(hwnd);
                }
            }
        }
    }

    fn adjust_transparency(&mut self, delta: i8) {
        self.transparency = (self.transparency as i16 + delta as i16)
            .clamp(50, 255) as u8;
        
        self.apply_all_windows(false);
        
        for (_, hwnd) in &self.floating_windows {
            unsafe {
                if IsWindow(*hwnd).as_bool() {
                    let ex_style = WINDOW_EX_STYLE(GetWindowLongW(*hwnd, GWL_EXSTYLE) as u32);
                    SetWindowLongW(*hwnd, GWL_EXSTYLE, 
                        (ex_style.0 | WS_EX_LAYERED.0) as i32);
                    SetLayeredWindowAttributes(*hwnd, COLORREF(0), self.transparency, LWA_ALPHA).ok();
                }
            }
        }
    }

    fn adjust_margins(&mut self, delta: i32) {
        self.margin_horizontal = (self.margin_horizontal as i32 + delta).clamp(0, 200) as i32;
        self.margin_vertical = (self.margin_vertical as i32 + delta * 2).clamp(0, 200) as i32;
        
        self.apply_all_windows(false);
    }
    
    fn cycle_fps(&mut self) {
        self.animation_fps = match self.animation_fps {
            60 => 90,
            90 => 120,
            120 => 144,
            _ => 60,
        };
    }
    
    fn scroll_to_window(&mut self, hwnd: HWND) {
        self.check_monitor_dimensions();
        
        if let Some(window) = self.windows.get(&hwnd.0) {
            // Extract values before mutable operations
            let window_row = window.position.row;
            let window_x = window.position.x;
            let window_size = window.position.size;
            
            // Set both vertical and horizontal targets
            self.current_row = window_row;
            self.vertical_offset_target = window_row * self.row_height;
            
            // Center the window horizontally
            let window_width = self.get_tile_width(&window_size);
            let center_offset = window_x + window_width / 2 - self.monitor_width / 2;
            
            let max_x = self.windows.values()
                .map(|w| w.position.x + self.get_tile_width(&w.position.size))
                .max()
                .unwrap_or(0);
            let max_offset = (max_x - self.monitor_width).max(0);
            
            self.ribbon_offset_target = center_offset.clamp(0, max_offset);
            
            // Start animation to both targets
            self.start_scroll_animation();
        }
    }
}

#[derive(Debug, Copy, Clone, PartialEq)]
enum Direction {
    Up,
    Down,
    Left,
    Right,
}

// Global state
static TILER: Mutex<Option<Arc<Mutex<RibbonTiler>>>> = Mutex::new(None);
static MAIN_HWND: AtomicUsize = AtomicUsize::new(0);
static SHUTDOWN_REQUESTED: AtomicBool = AtomicBool::new(false);

// Keyboard hook procedure
unsafe extern "system" fn keyboard_hook_proc(
    code: i32,
    wparam: WPARAM,
    lparam: LPARAM,
) -> LRESULT {
    if code < 0 {
        return CallNextHookEx(HHOOK::default(), code, wparam, lparam);
    }

    if wparam.0 as u32 == WM_KEYDOWN || wparam.0 as u32 == WM_SYSKEYDOWN {
        let kb_struct = *(lparam.0 as *const KBDLLHOOKSTRUCT);
        let vk_code = VIRTUAL_KEY(kb_struct.vkCode as u16);
        
        if kb_struct.flags.contains(KBDLLHOOKSTRUCT_FLAGS(0x10)) {
            return CallNextHookEx(HHOOK::default(), code, wparam, lparam);
        }
        
        let ctrl = GetAsyncKeyState(VK_CONTROL.0 as i32) & 0x8000u16 as i16 != 0;
        let alt = GetAsyncKeyState(VK_MENU.0 as i32) & 0x8000u16 as i16 != 0;
        let win = GetAsyncKeyState(VK_LWIN.0 as i32) & 0x8000u16 as i16 != 0 
            || GetAsyncKeyState(VK_RWIN.0 as i32) & 0x8000u16 as i16 != 0;
        let shift = GetAsyncKeyState(VK_SHIFT.0 as i32) & 0x8000u16 as i16 != 0;

        if !win {
            return CallNextHookEx(HHOOK::default(), code, wparam, lparam);
        }

        let main_hwnd_value = MAIN_HWND.load(Ordering::Relaxed);
        if main_hwnd_value == 0 {
            return CallNextHookEx(HHOOK::default(), code, wparam, lparam);
        }
        let main_hwnd = HWND(main_hwnd_value as isize);
        
        let hwnd = GetForegroundWindow();
        
        let mut command: Option<TilerCommand> = None;
        
        if win && !ctrl && !shift && !alt {
            match vk_code {
                VK_UP => command = Some(TilerCommand::PanUp),
                VK_DOWN => command = Some(TilerCommand::PanDown),
                VK_LEFT => command = Some(TilerCommand::PanLeft),
                VK_RIGHT => command = Some(TilerCommand::PanRight),
                VIRTUAL_KEY(0x43) => command = Some(TilerCommand::ForceRecalc), // C for Clean
                _ => {},
            }
        }

        if win && ctrl && !shift && !alt {
            match vk_code {
                VK_LEFT => command = Some(TilerCommand::ResizeLeft),
                VK_RIGHT => command = Some(TilerCommand::ResizeRight),
                _ => {},
            }
        }

        if win && ctrl && shift && !alt {
            match vk_code {
                VK_UP => command = Some(TilerCommand::MoveUp),
                VK_DOWN => command = Some(TilerCommand::MoveDown),
                VK_LEFT => command = Some(TilerCommand::MoveLeft),
                VK_RIGHT => command = Some(TilerCommand::MoveRight),
                _ => {},
            }
        }
        
        if win && !ctrl && !shift && !alt {
            match vk_code {
                VK_OEM_PLUS | VK_ADD => command = Some(TilerCommand::IncreaseTransparency),
                VK_OEM_MINUS | VK_SUBTRACT => command = Some(TilerCommand::DecreaseTransparency),
                VIRTUAL_KEY(0x53) => command = Some(TilerCommand::ScrollToWindow), // S
                VIRTUAL_KEY(0x4D) => command = Some(TilerCommand::IncreaseMargins), // M
                VIRTUAL_KEY(0x4E) => command = Some(TilerCommand::DecreaseMargins), // N
                VIRTUAL_KEY(0x46) => command = Some(TilerCommand::CycleFPS), // F for FPS
                _ => {},
            }
        }

        if win && shift && !ctrl && !alt {
            match vk_code {
                VK_OEM_PLUS | VK_ADD => command = Some(TilerCommand::IncreaseTransparency),
                VK_OEM_MINUS | VK_SUBTRACT => command = Some(TilerCommand::DecreaseTransparency),
                VIRTUAL_KEY(0x54) => command = Some(TilerCommand::AddWindow), // T
                VIRTUAL_KEY(0x52) => command = Some(TilerCommand::RemoveWindow), // R
                _ => {},
            }
        }
        
        if let Some(cmd) = command {
            PostMessageW(
                main_hwnd,
                WM_TILER_COMMAND,
                WPARAM(cmd as usize),
                LPARAM(hwnd.0)
            ).ok();
            return LRESULT(1);
        }
    }
    
    CallNextHookEx(HHOOK::default(), code, wparam, lparam)
}

// Handler for Ctrl+C signal
extern "system" fn console_handler(ctrl_type: u32) -> BOOL {
    const CTRL_C_EVENT: u32 = 0;
    if ctrl_type == CTRL_C_EVENT {
        SHUTDOWN_REQUESTED.store(true, Ordering::Relaxed);
        
        let main_hwnd_value = MAIN_HWND.load(Ordering::Relaxed);
        if main_hwnd_value != 0 {
            unsafe {
                PostMessageW(
                    HWND(main_hwnd_value as isize),
                    WM_TILER_SHUTDOWN,
                    WPARAM(0),
                    LPARAM(0)
                ).ok();
            }
        }
        
        BOOL::from(true)
    } else {
        BOOL::from(false)
    }
}

fn main() -> Result<()> {
    println!("╔═══════════════════════════════════════════════╗");
    println!("║     THYMELINE TILER v3.1 - Smooth Scrolling    ║");
    println!("╚═══════════════════════════════════════════════╝");
    println!("\n🎯 WINDOW MANAGEMENT:");
    println!("  Win+Shift+T          Add current window to ribbon");
    println!("  Win+Shift+R          Remove current window from ribbon");
    println!("  Win+C                Force cleanup and recalculation");
    println!("\n📐 WINDOW RESIZING:");
    println!("  Win+Ctrl+Left/Right  Toggle between full/half width");
    println!("\n🔀 WINDOW MOVEMENT:");
    println!("  Win+Ctrl+Shift+Arrow Move windows (up/down changes rows)");
    println!("\n📍 RIBBON NAVIGATION:");
    println!("  Win+Left/Right       Pan horizontally through ribbon");
    println!("  Win+Up/Down          Switch between rows");
    println!("  Win+S                Scroll to current window");
    println!("\n🎨 APPEARANCE:");
    println!("  Win+Plus             Increase transparency");
    println!("  Win+Minus            Decrease transparency");
    println!("  Win+M                Increase margins (+5H/+10V)");
    println!("  Win+N                Decrease margins (-5H/-10V)");
    println!("  Win+F                Cycle FPS (60→90→120→144)");
    println!("\nPress Ctrl+C to exit gracefully");

    unsafe {
        if SetConsoleCtrlHandler(Some(console_handler), true).is_err() {
            println!("Warning: Failed to set console handler");
        }
        
        let tiler = Arc::new(Mutex::new(RibbonTiler::new()));
        
        {
            let tiler_lock = tiler.lock().unwrap();
            MAIN_HWND.store(tiler_lock.main_hwnd.0 as usize, Ordering::Relaxed);
        }
        
        *TILER.lock().unwrap() = Some(tiler.clone());

        let hook = SetWindowsHookExW(
            WH_KEYBOARD_LL,
            Some(keyboard_hook_proc),
            GetModuleHandleW(None)?,
            0,
        )?;

        let mut msg = MSG::default();
        loop {
            let result = GetMessageW(&mut msg, HWND::default(), 0, 0);
            if result.0 == 0 || result.0 == -1 {
                break;
            }
            
            if msg.message == WM_USER + 1 {
                if let Some(tiler_arc) = TILER.lock().unwrap().as_ref() {
                    if let Ok(mut tiler) = tiler_arc.lock() {
                        tiler.update_animations();
                    }
                }
            } else if msg.message == WM_TILER_RECALC {
                if let Some(tiler_arc) = TILER.lock().unwrap().as_ref() {
                    if let Ok(mut tiler) = tiler_arc.lock() {
                        tiler.clean_closed_windows();
                        
                        if tiler.needs_ribbon_recalc {
                            let now = Instant::now();
                            if now.duration_since(tiler.last_ribbon_recalc).as_millis() > 500 {
                                tiler.recalculate_ribbon();
                            }
                        }
                    }
                }
            } else if msg.message == WM_TILER_COMMAND {
                if let Some(tiler_arc) = TILER.lock().unwrap().as_ref() {
                    if let Ok(mut tiler) = tiler_arc.lock() {
                        let command_value = msg.wParam.0 as u32;
                        let hwnd = HWND(msg.lParam.0);
                        
                        let command = match command_value {
                            0 => TilerCommand::PanLeft,
                            1 => TilerCommand::PanRight,
                            2 => TilerCommand::PanUp,
                            3 => TilerCommand::PanDown,
                            4 => TilerCommand::ResizeLeft,
                            5 => TilerCommand::ResizeRight,
                            6 => TilerCommand::MoveUp,
                            7 => TilerCommand::MoveDown,
                            8 => TilerCommand::MoveLeft,
                            9 => TilerCommand::MoveRight,
                            10 => TilerCommand::AddWindow,
                            14 => TilerCommand::IncreaseTransparency,
                            15 => TilerCommand::DecreaseTransparency,
                            17 => TilerCommand::ScrollToWindow,
                            18 => TilerCommand::IncreaseMargins,
                            19 => TilerCommand::DecreaseMargins,
                            20 => TilerCommand::RemoveWindow,
                            21 => TilerCommand::CycleFPS,
                            22 => TilerCommand::ForceRecalc,
                            _ => continue,
                        };
                        
                        match command {
                            TilerCommand::PanLeft | TilerCommand::PanRight | 
                            TilerCommand::PanUp | TilerCommand::PanDown => {
                                // Process pan commands immediately for smooth aggregation
                                match command {
                                    TilerCommand::PanLeft => tiler.pan_ribbon(Direction::Left),
                                    TilerCommand::PanRight => tiler.pan_ribbon(Direction::Right),
                                    TilerCommand::PanUp => tiler.pan_row(Direction::Up),
                                    TilerCommand::PanDown => tiler.pan_row(Direction::Down),
                                    _ => {},
                                }
                            },
                            _ => {
                                tiler.queue_command(command, hwnd);
                                tiler.process_command_queue();
                            }
                        }
                    }
                }
            } else if msg.message == WM_TILER_SHUTDOWN {
                break;
            }
            
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }

        if let Some(tiler_arc) = TILER.lock().unwrap().as_ref() {
            if let Ok(mut tiler) = tiler_arc.lock() {
                tiler.shutdown();
            }
        }
        
        UnhookWindowsHookEx(hook)?;
        println!("\nThymeline shut down gracefully");
    }
    Ok(())
}
