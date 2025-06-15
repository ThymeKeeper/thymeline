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
    ResizeUp = 2,
    ResizeDown = 3,
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

// Window size variants
#[derive(Debug, Clone, Copy, PartialEq)]
enum TileSize {
    Full,           // Full screen
    HalfHorizontal, // Horizontal split - full width, half height (short and wide)
    HalfVertical,   // Vertical split - half width, full height (tall and narrow)
    Quarter,        // Quarter screen
}

// Position in the ribbon (x is the virtual position)
#[derive(Debug, Clone, Copy)]
struct RibbonPosition {
    x: i32,  // Virtual x position in ribbon
    y: i32,  // Y position (0 or 1 for top/bottom half)
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
    ribbon_offset_animation: Option<(Instant, i32, i32)>,
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
    main_hwnd: HWND, // Hidden window for message processing
    command_queue: Vec<QueuedCommand>,
    last_command_time: HashMap<u32, Instant>, // Track last execution time per command type
    animation_fps: u64, // Configurable animation frame rate
    needs_ribbon_recalc: bool, // Flag to indicate ribbon needs recalculation
    last_ribbon_recalc: Instant, // Last time ribbon was recalculated
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
        
        //println!("Ribbon Tiler initialized - Monitor: {}x{}, Margins: H:{}px V:{}px, Animation: {} FPS", width, height, 40, 80, 90);
        Self {
            windows: HashMap::new(),
            floating_windows: HashMap::new(),
            ribbon_offset: 0,
            ribbon_offset_target: 0,
            ribbon_offset_animation: None,
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
                TilerCommand::PanLeft | TilerCommand::PanRight => false,
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
                TilerCommand::ResizeUp => self.resize_window(queued.hwnd, Direction::Up),
                TilerCommand::ResizeDown => self.resize_window(queued.hwnd, Direction::Down),
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
                    // Immediately trigger recalculation instead of waiting
                    if self.needs_ribbon_recalc {
                        self.recalculate_ribbon();
                    }
                },
                TilerCommand::CycleFPS => self.cycle_fps(),
                TilerCommand::ForceRecalc => {
                    //println!("\nForcing cleanup and recalculation...");
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
        let window_start = pos.x - self.ribbon_offset;
        let window_end = window_start + self.get_tile_width(&pos.size);
        let buffer = self.monitor_width;
        window_end >= -buffer && window_start <= self.monitor_width + buffer
    }

    // Update animations - OPTIMIZED VERSION with batching and visibility culling
    fn update_animations(&mut self) {
        let now = Instant::now();
        let mut animations_complete = Vec::new();
        let mut window_updates = Vec::new();
        let mut need_reposition = false;

        // Update ribbon offset animation
        if let Some((start_time, start_offset, target_offset)) = self.ribbon_offset_animation {
            let elapsed = now.duration_since(start_time);
            let duration = Duration::from_millis(87);
            
            if elapsed >= duration {
                self.ribbon_offset = target_offset;
                self.ribbon_offset_animation = None;
                self.focus_visible_window();
                self.needs_ribbon_recalc = true;
            } else {
                let t = elapsed.as_secs_f32() / duration.as_secs_f32();
                let eased_t = Self::ease_out_cubic(t);
                self.ribbon_offset = Self::lerp(start_offset, target_offset, eased_t);
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
        let all_complete = self.ribbon_offset_animation.is_none() &&
            self.windows.values().all(|w| w.animation.is_none());
        
        if all_complete {
            *self.animation_running.lock().unwrap() = false;
            *self.animation_stop_requested.lock().unwrap() = true;
            
            // Check if ribbon needs recalculation (with debounce)
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
                        
                        if width > 0 && height > 0 && 
                           rect.left > -10000 && rect.top > -10000 && 
                           rect.right < 10000 && rect.bottom < 10000 {
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
                        
                        if width > 0 && height > 0 {
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
            
            if rect.left < -10000 || rect.top < -10000 || rect.right > 10000 || rect.bottom > 10000 {
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
                            
                            // Check if we need to trigger a ribbon recalc (every 500ms)
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
        
        let (w, h) = match pos.size {
            TileSize::Full => (self.monitor_width, self.monitor_height),
            TileSize::HalfHorizontal => (self.monitor_width, self.monitor_height / 2),
            TileSize::HalfVertical => (self.monitor_width / 2, self.monitor_height),
            TileSize::Quarter => (self.monitor_width / 2, self.monitor_height / 2),
        };

        let y = match (pos.size, pos.y) {
            (TileSize::HalfHorizontal, 1) => self.monitor_height / 2,
            (TileSize::Quarter, 1) => self.monitor_height / 2,
            _ => 0,
        };

        RECT {
            left: base_x + self.margin_horizontal / 2,
            top: y + self.margin_vertical / 2,
            right: base_x + w - self.margin_horizontal / 2,
            bottom: y + h - self.margin_vertical / 2,
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

            println!("Window added to ribbon (pushed in at nearest edge)");
            
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

    // NEW: Clean up windows that were closed externally
    fn clean_closed_windows(&mut self) {
        let mut closed_windows = Vec::new();
        
        unsafe {
            for (hwnd_val, _) in self.windows.iter() {
                let hwnd = HWND(*hwnd_val);
                // Check if window still exists and is visible
                if !IsWindow(hwnd).as_bool() || !IsWindowVisible(hwnd).as_bool() {
                    closed_windows.push(*hwnd_val);
                }
            }
        }
        
        if !closed_windows.is_empty() {
            //println!("\nDetected {} externally closed windows, removing from ribbon", closed_windows.len());
            for hwnd_val in &closed_windows {
                self.windows.remove(hwnd_val);
            }
            self.needs_ribbon_recalc = true;
        }
    }

    // NEW: Recalculate entire ribbon layout
    fn recalculate_ribbon(&mut self) {
        // First clean up any windows that were closed externally
        self.clean_closed_windows();
        
        if self.windows.is_empty() {
            self.ribbon_offset = 0;
            self.ribbon_offset_target = 0;
            //println!("\nRibbon is empty - no windows to manage");
            return;
        }
        
        //println!("\nRecalculating ribbon layout...");
        
        // Get all non-exiting windows sorted by current x position
        let mut windows: Vec<(isize, RibbonPosition)> = self.windows.iter()
            .filter(|(_, w)| {
                w.animation.as_ref().map_or(true, |a| a.animation_type != AnimationType::Exit)
            })
            .map(|(hwnd, w)| (*hwnd, w.position))
            .collect();
        
        //println!("Found {} active windows in ribbon", windows.len());
        
        // Sort by the current x position to maintain relative order
        windows.sort_by_key(|(_, pos)| pos.x);
        
        // Debug: print current positions before recalc
        //println!("Current window positions:");
        //for (i, (hwnd, pos)) in windows.iter().enumerate() {
        //    println!("  Window {}: x={}, width={}", i, pos.x, self.get_tile_width(&pos.size));
        //}
        
        // Reposition all windows starting from x=0, closing all gaps
        let mut current_x = 0;
        let mut positions_to_update = Vec::new();
        let mut position_changes = Vec::new();
        
        for (hwnd, old_position) in windows {
            let width = self.get_tile_width(&old_position.size);
            
            if let Some(window) = self.windows.get_mut(&hwnd) {
                let old_x = window.position.x;
                
                // Update position to close gaps
                window.position.x = current_x;
                // Keep y and size from the old position
                window.position.y = old_position.y;
                window.position.size = old_position.size;
                
                if old_x != current_x {
                    position_changes.push((hwnd, old_x, current_x));
                }
                
                // Store windows that need visual position updates
                if window.animation.is_none() {
                    positions_to_update.push((window.hwnd, window.position));
                }
            }
            
            current_x += width;
        }
        
        let total_width = current_x;
        
        // Apply position updates after mutable borrows are done
        for (hwnd, position) in positions_to_update {
            let rect = self.ribbon_to_screen(&position);
            Self::set_window_rect(hwnd, &rect);
        }
        
        // Ensure ribbon offset is within bounds
        let max_offset = (total_width - self.monitor_width).max(0);
        if self.ribbon_offset > max_offset {
            self.ribbon_offset = max_offset;
            self.ribbon_offset_target = max_offset;
            self.ribbon_offset_animation = None;
        }
        
        self.needs_ribbon_recalc = false;
        self.last_ribbon_recalc = Instant::now();
    }
    
    // Clamp ribbon offset to valid bounds (simplified - just triggers recalc)
    fn clamp_ribbon_offset(&mut self) {
        self.needs_ribbon_recalc = true;
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
            
            // Remove minimize/maximize functionality
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
            
            let windows_to_shift: Vec<isize> = self.windows.iter()
                .filter(|(h, w)| **h != hwnd.0 && w.position.x >= insertion_x)
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
            }
            
            self.apply_window_position_with_animation_type(hwnd, AnimationType::Entry);
            
            let shifted_hwnds: Vec<HWND> = self.windows.iter()
                .filter(|(h, w)| **h != hwnd.0 && w.position.x >= insertion_x + new_window_width)
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
        
        for window in self.windows.values() {
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
            y: 0,
            size: TileSize::HalfVertical,
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
            y: 0,
            size: TileSize::HalfVertical,
        }
    }

    fn get_tile_width(&self, size: &TileSize) -> i32 {
        match size {
            TileSize::Full => self.monitor_width,
            TileSize::HalfHorizontal => self.monitor_width,
            TileSize::HalfVertical => self.monitor_width / 2,
            TileSize::Quarter => self.monitor_width / 2,
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
            let old_y = window.position.y;
            
            let (new_size, new_y) = match (old_size, old_y, direction) {
                // From Full
                (TileSize::Full, _, Direction::Left | Direction::Right) => (TileSize::HalfVertical, 0),
                (TileSize::Full, _, Direction::Up) => (TileSize::HalfHorizontal, 0),
                (TileSize::Full, _, Direction::Down) => (TileSize::HalfHorizontal, 1),
                
                // From HalfHorizontal
                (TileSize::HalfHorizontal, 0, Direction::Up) => (TileSize::HalfHorizontal, 0),
                (TileSize::HalfHorizontal, 0, Direction::Down) => (TileSize::Full, 0),
                (TileSize::HalfHorizontal, 1, Direction::Up) => (TileSize::Full, 0),
                (TileSize::HalfHorizontal, 1, Direction::Down) => (TileSize::HalfHorizontal, 1),
                (TileSize::HalfHorizontal, y, Direction::Left | Direction::Right) => (TileSize::Quarter, y),
                
                // From HalfVertical
                (TileSize::HalfVertical, _, Direction::Up) => (TileSize::Quarter, 0),
                (TileSize::HalfVertical, _, Direction::Down) => (TileSize::Quarter, 1),
                (TileSize::HalfVertical, _, Direction::Left | Direction::Right) => (TileSize::Full, 0),
                
                // From Quarter
                (TileSize::Quarter, 0, Direction::Up) => (TileSize::Quarter, 0),
                (TileSize::Quarter, 0, Direction::Down) => (TileSize::HalfVertical, 0),
                (TileSize::Quarter, 1, Direction::Up) => (TileSize::HalfVertical, 0),
                (TileSize::Quarter, 1, Direction::Down) => (TileSize::Quarter, 1),
                (TileSize::Quarter, y, Direction::Left | Direction::Right) => (TileSize::HalfHorizontal, y),
                
                _ => (old_size, old_y),
            };
            
            if let Some(w) = self.windows.get_mut(&hwnd.0) {
                w.position.size = new_size;
                w.position.y = new_y;
            }
            
            self.pull_adjacent_windows(hwnd.0);
            self.needs_ribbon_recalc = true;
        }
        
        self.apply_all_windows(true);
    }

    fn pull_adjacent_windows(&mut self, _changed_hwnd: isize) {
        self.needs_ribbon_recalc = true;
    }

    fn move_window(&mut self, hwnd: HWND, direction: Direction) {
        self.check_monitor_dimensions();
        self.clean_closed_windows();
        self.clean_minimized_windows();
        
        if !self.windows.contains_key(&hwnd.0) {
            return;
        }

        let current_pos = match self.windows.get(&hwnd.0) {
            Some(w) => w.position,
            None => return,
        };
        
        match direction {
            Direction::Up | Direction::Down => {
                if current_pos.size == TileSize::Quarter {
                    if let Some(window) = self.windows.get_mut(&hwnd.0) {
                        window.position.y = 1 - window.position.y;
                    }
                    self.apply_all_windows(true);
                }
            },
            Direction::Left | Direction::Right => {
                if let Some(window) = self.windows.get_mut(&hwnd.0) {
                    window.animation = None;
                }
                
                let focused_screen_x = current_pos.x - self.ribbon_offset;
                
                let mut windows: Vec<(isize, i32)> = self.windows.iter()
                    .map(|(h, w)| (*h, w.position.x))
                    .collect();
                windows.sort_by_key(|(_, x)| *x);
                
                let current_index = windows.iter().position(|(h, _)| *h == hwnd.0);
                if current_index.is_none() {
                    return;
                }
                let current_index = current_index.unwrap();
                
                match direction {
                    Direction::Left if current_index == 0 => return,
                    Direction::Right if current_index == windows.len() - 1 => return,
                    _ => {}
                }
                
                match direction {
                    Direction::Left => {
                        let (left_hwnd, left_x) = windows[current_index - 1];
                        let (_current_hwnd, current_x) = windows[current_index];
                        
                        if let Some(left_window) = self.windows.get_mut(&left_hwnd) {
                            left_window.position.x = current_x;
                        }
                        if let Some(current_window) = self.windows.get_mut(&hwnd.0) {
                            current_window.position.x = left_x;
                        }
                        
                        self.ribbon_offset = left_x - focused_screen_x;
                        self.needs_ribbon_recalc = true;
                    },
                    Direction::Right => {
                        let (_current_hwnd, current_x) = windows[current_index];
                        let (right_hwnd, right_x) = windows[current_index + 1];
                        
                        if let Some(right_window) = self.windows.get_mut(&right_hwnd) {
                            right_window.position.x = current_x;
                        }
                        if let Some(current_window) = self.windows.get_mut(&hwnd.0) {
                            current_window.position.x = right_x;
                        }
                        
                        self.ribbon_offset = right_x - focused_screen_x;
                        self.needs_ribbon_recalc = true;
                    },
                    _ => unreachable!(),
                }
                
                let max_x = self.windows.values()
                    .map(|w| w.position.x + self.get_tile_width(&w.position.size))
                    .max()
                    .unwrap_or(0);
                let max_offset = (max_x - self.monitor_width).max(0);
                
                self.ribbon_offset = self.ribbon_offset.clamp(0, max_offset);
                self.ribbon_offset_target = self.ribbon_offset;
                
                self.ribbon_offset_animation = None;
                
                if let Some(window) = self.windows.get(&hwnd.0) {
                    let rect = self.ribbon_to_screen(&window.position);
                    Self::set_window_rect(hwnd, &rect);
                }
                
                let other_hwnds: Vec<HWND> = self.windows.iter()
                    .filter(|(h, _)| **h != hwnd.0)
                    .map(|(_, w)| w.hwnd)
                    .collect();
                    
                for other_hwnd in other_hwnds {
                    self.apply_window_position(other_hwnd, true);
                }
                
                self.start_animation_timer();
                
                unsafe {
                    SetForegroundWindow(hwnd);
                }
            }
        }
    }

    fn recalculate_positions_for_new_resolution(&mut self, old_width: i32) {
        if self.windows.is_empty() {
            return;
        }
        
        let mut windows: Vec<(isize, RibbonPosition)> = self.windows.iter()
            .map(|(hwnd, w)| (*hwnd, w.position))
            .collect();
        windows.sort_by_key(|(_, pos)| pos.x);
        
        let scale_factor = self.monitor_width as f32 / old_width as f32;
        
        let mut current_x = 0;
        for (hwnd, old_pos) in windows {
            if let Some(window) = self.windows.get_mut(&hwnd) {
                window.position.x = current_x;
                window.position.y = old_pos.y;
                window.position.size = old_pos.size;
                
                current_x += self.get_tile_width(&old_pos.size);
            }
        }
        
        self.ribbon_offset = (self.ribbon_offset as f32 * scale_factor) as i32;
        self.ribbon_offset_target = self.ribbon_offset;
        
        self.needs_ribbon_recalc = true;
    }

    fn apply_all_windows(&mut self, animate: bool) {
        let hwnds: Vec<HWND> = self.windows.values().map(|w| w.hwnd).collect();
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
            
            //println!("Monitor resolution changed: {}x{} -> {}x{}", old_width, self.monitor_height, new_width, new_height);
            
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
        
        match direction {
            Direction::Left if self.ribbon_offset_target <= 0 => return,
            Direction::Right if self.ribbon_offset_target >= max_offset => return,
            _ => {}
        }
        
        let snap_distance = self.monitor_width / 2;
        let old_target = self.ribbon_offset_target;
        
        match direction {
            Direction::Left => {
                self.ribbon_offset_target = (self.ribbon_offset_target - snap_distance).max(0);
            },
            Direction::Right => {
                self.ribbon_offset_target = (self.ribbon_offset_target + snap_distance).min(max_offset);
            },
            _ => return,
        }
        
        if self.ribbon_offset_target != old_target {
            self.ribbon_offset_animation = Some((
                Instant::now(),
                self.ribbon_offset,
                self.ribbon_offset_target
            ));
            
            self.start_animation_timer();
            self.focus_visible_window();
        }
    }
    
    fn focus_visible_window(&self) {
        unsafe {
            let target_offset = self.ribbon_offset_target;
            
            let mut best_window: Option<HWND> = None;
            let mut best_distance = i32::MAX;
            
            for window in self.windows.values() {
                let window_start = window.position.x;
                let window_width = self.get_tile_width(&window.position.size);
                let window_end = window_start + window_width;
                let window_center = window_start + window_width / 2;
                
                if window_end > target_offset && window_start < target_offset + self.monitor_width {
                    let screen_center = target_offset + self.monitor_width / 2;
                    let distance = (window_center - screen_center).abs();
                    
                    if distance < best_distance {
                        best_distance = distance;
                        best_window = Some(window.hwnd);
                    }
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
        //println!("Window margins: H:{}px V:{}px", self.margin_horizontal, self.margin_vertical);
        
        self.apply_all_windows(false);
    }
    
    fn cycle_fps(&mut self) {
        self.animation_fps = match self.animation_fps {
            60 => 90,
            90 => 120,
            120 => 144,
            _ => 60,
        };
        //println!("Animation FPS: {} ({}ms frame interval)", self.animation_fps, 1000 / self.animation_fps);
    }
    
    fn scroll_to_window(&mut self, hwnd: HWND) {
        self.check_monitor_dimensions();
        
        if let Some(window) = self.windows.get(&hwnd.0) {
            let window_x = window.position.x;
            let window_width = self.get_tile_width(&window.position.size);
            let window_end = window_x + window_width;
            
            if window_x >= self.ribbon_offset && window_end <= self.ribbon_offset + self.monitor_width {
                return;
            }
            
            let center_offset = window_x + window_width / 2 - self.monitor_width / 2;
            
            let max_x = self.windows.values()
                .map(|w| w.position.x + self.get_tile_width(&w.position.size))
                .max()
                .unwrap_or(0);
            let max_offset = (max_x - self.monitor_width).max(0);
            
            self.ribbon_offset_target = center_offset.clamp(0, max_offset);
            
            self.ribbon_offset_animation = Some((
                Instant::now(),
                self.ribbon_offset,
                self.ribbon_offset_target
            ));
            
            self.start_animation_timer();
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

// Keyboard hook procedure - MUST BE FAST!
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
                VK_UP | VK_DOWN => return LRESULT(1),
                _ => {},
            }
        }
        
        if win && !ctrl && !shift && !alt {
            match vk_code {
                VK_LEFT => command = Some(TilerCommand::PanLeft),
                VK_RIGHT => command = Some(TilerCommand::PanRight),
                VIRTUAL_KEY(0x43) => command = Some(TilerCommand::ForceRecalc), // C for Clean
                _ => {},
            }
        }

        if win && ctrl && !shift && !alt {
            match vk_code {
                VK_UP => command = Some(TilerCommand::ResizeUp),
                VK_DOWN => command = Some(TilerCommand::ResizeDown),
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
    println!("║             THYMELINE TILER v2.7.1            ║");
    println!("╚═══════════════════════════════════════════════╝");
    println!("\n🎯 WINDOW MANAGEMENT:");
    println!("  Win+Shift+T          Add current window to ribbon");
    println!("  Win+Shift+R          Remove current window from ribbon");
    println!("  Win+C                Force cleanup and recalculation");
    println!("\n📐 WINDOW RESIZING:");
    println!("  Win+Ctrl+Arrow       Resize window");
    println!("  Win+Up/Down          (Disabled - no maximize/minimize)");
    println!("\n🔀 WINDOW MOVEMENT:");
    println!("  Win+Ctrl+Shift+Arrow Move/shuffle windows");
    println!("\n📍 RIBBON NAVIGATION:");
    println!("  Win+Left/Right       Pan through ribbon");
    println!("  Win+S                Scroll to current window");
    println!("\n🎨 APPEARANCE:");
    println!("  Win+(Shift+)Plus     Increase transparency");
    println!("  Win+(Shift+)Minus    Decrease transparency");
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

        //println!("\n✅ Thymeline ready! (v2.7.1 - Smart ribbon with closed window detection)");

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
                        // Always clean closed windows when we check for recalc
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
                            2 => TilerCommand::ResizeUp,
                            3 => TilerCommand::ResizeDown,
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
                            TilerCommand::PanLeft => tiler.pan_ribbon(Direction::Left),
                            TilerCommand::PanRight => tiler.pan_ribbon(Direction::Right),
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