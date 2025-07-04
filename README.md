# WindowPots

A window organizer for Windows that arranges application windows in a scrollable 2D grid with smooth animations.

![thymeline demo vid](https://github.com/user-attachments/assets/2cff35d6-beba-4fc5-92e3-7b95d1f1fd0b)

## What is WindowPots?

WindowPots is a Windows application that manages other application windows, arranging them in a 2D plane where you can scroll horizontally through columns and vertically through rows. Each window becomes a tile in this grid.

**Note:** This is not a true window manager - it's an application that repositions and manages other windows. Your regular Windows desktop environment remains unchanged.. it was designed this way to help users stay organized on windows machines where they don't have admin privileges.

## Features

- **2D Grid Layout** - Tiles arranged in rows and columns
- **Smooth Scrolling** - Navigate horizontally and vertically through your tile grid
- **Smart Positioning** - Tiles automatically arrange themselves without gaps
- **Adjustable Transparency** - Set tile transparency (50-255 alpha)
- **Dynamic Margins** - Adjust spacing between tiles
- **Variable Frame Rates** - 60/90/120/144 FPS animation options
- **Popup Handling** - Dialog boxes and popups remain floating
- **Entry/Exit Animations** - Visual feedback when adding/removing tiles

## Keyboard Shortcuts

### Window Management
| Shortcut | Action |
|----------|--------|
| `Win+Shift+T` | Add current floating window to the grid |
| `Win+Shift+R` | Remove current tile from grid |
| `Win+C` | Force cleanup and recalculation |

### Window Sizing
| Shortcut | Action |
|----------|--------|
| `Win+Ctrl+←/→` | Toggle tile between full/half width |

### Window Movement
| Shortcut | Action |
|----------|--------|
| `Win+Ctrl+Shift+arrow` | Move focused tile |

### Navigation
| Shortcut | Action |
|----------|--------|
| `Win+arrow` | Pan view |
| `Win+S` | Scroll to focused window |

### Appearance
| Shortcut | Action |
|----------|--------|
| `Win+[=]` | Increase transparency |
| `Win+[-]` | Decrease transparency |
| `Win+M` | Increase margins |
| `Win+N` | Decrease margins |
| `Win+F` | Cycle animation FPS |

### Exit
| Shortcut | Action |
|----------|--------|
| `Ctrl+C` | Restore all windows and exit (when the main terminal window is focused) |

### Prerequisites
- Rust
- Windows 10/11

## How It Works

1. **Start WindowPots** - Run the executable
2. **Add Windows** - Focus any window and press `Win+Shift+T` to add it to the grid
3. **Navigate** - Use `Win+Arrow` keys to move through the 2D plane
4. **Organize** - Reposition and resize windows as needed
5. **Exit Cleanly** - Press `Ctrl+C` to restore all windows to their original positions

## Known Limitations

- Only works on Windows 10/11
- Some applications may not respond well to window manipulation
- UWP/Modern apps might have limited functionality
- Multiple monitor setups haven't been tested at all; this was designed for a single monitor

## Personal Project Notice

This is a personal project built for my own use. Development is driven by my needs and interests, but feel free to reach out if it interests you.

---

*Built with Rust and the windows-rs crate*
