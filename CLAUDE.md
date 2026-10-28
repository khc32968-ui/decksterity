# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**Decksterity** is a Microsoft PowerPoint VSTO (Visual Studio Tools for Office) add-in that provides enhanced slide elements and layout tools. The add-in creates a custom ribbon with two tabs: "Decksterity" and "Decksterity Expert", offering a comprehensive suite of visual elements and formatting tools.

## Technology Stack

- **Framework**: .NET Framework 4.7.2
- **Platform**: C# VSTO Add-in for Microsoft PowerPoint
- **Office Integration**: Microsoft Office Interop APIs
- **Project Type**: Visual Studio Office Add-in Project

## Architecture

### Core Components

1. **ThisAddIn.cs**: Main add-in entry point that registers the ribbon extensibility and manages keyboard hooks
2. **DecksterityRibbon.cs**: Implements IRibbonExtensibility interface, handles ribbon callbacks and image loading
3. **DecksterityRibbon.xml**: Defines the ribbon UI structure with tabs, groups, and buttons
4. **ElementHelper.cs**: Centralized utility class for inserting all types of symbols into PowerPoint slides
5. **AlignmentHelper.cs**: Comprehensive alignment, distribution, and spacing utilities for PowerPoint shapes
6. **KeyboardHookManager.cs**: Windows keyboard hook implementation for Ctrl+Shift shortcut handling
7. **SlideElementHelper.cs**: Inserts consulting slide elements (sticky, table, header, quote, status sticker, footnote) for Ctrl+Shift+letter shortcuts
8. **ShapeToolsHelper.cs**: Arrange in grid, copy/paste position and size, and text tools (zero margins, autofit off, swap text, merge text boxes); each operation starts a new undo entry

### Key Features

- **Harvey Balls**: Unicode-based progress indicators (0-4 fill levels: ⭘, ◔, ◑, ◕, ●)
- **Arrows**: 8-directional arrow symbols (🡹, 🡽, 🡺, 🡾, 🡻, 🡿, 🡸, 🡼)
- **Icons**: Basic symbols (✔, ✘, ➕, ➖, ❓, …)
- **Stoplights**: Colored status indicators (red, amber, green solid circles)
- **Layout Tools**: Full alignment, distribution, and arrangement utilities
- **Keyboard Shortcuts**: Ctrl+Shift+1-6 shortcuts for common alignment operations

### File Structure

```
decksterity/
├── DecksterityRibbon.cs     # Main ribbon implementation
├── DecksterityRibbon.xml    # Ribbon UI definition
├── ElementHelper.cs         # Centralized element insertion logic
├── AlignmentHelper.cs       # Alignment, distribution, and spacing utilities
├── KeyboardHookManager.cs   # Windows keyboard hook for shortcuts
├── ThisAddIn.cs            # VSTO add-in entry point
├── assets/                 # PNG icons for ribbon buttons
│   └── generators/         # Jupyter notebooks for icon generation
├── Properties/             # Assembly info and resources
├── .github/workflows/      # GitHub Actions for automated deployment
│   └── release.yml         # MSI deployment workflow
├── README.md              # Public documentation and installation guide
└── CLAUDE.md              # Development guidance for Claude Code
```

## Development Commands

### Building the Project

```bash
# Build in Debug mode
msbuild decksterity.sln /p:Configuration=Debug

# Build in Release mode  
msbuild decksterity.sln /p:Configuration=Release
```

### Visual Studio Commands

- **F5**: Build and run with debugging (launches PowerPoint with add-in loaded)
- **Ctrl+Shift+B**: Build solution
- **F6**: Build current project

## PowerPoint Integration Patterns

### Smart Context-Aware Insertion
The ElementHelper class handles different PowerPoint selection contexts intelligently:

- **`ppSelectionText`**: Insert symbols at cursor position in existing text (with font preservation)
- **`ppSelectionShapes`**: 
  - Table cells: Insert into selected table cells
  - Text shapes: Replace/insert into shape text frames
- **Fallback**: Create new centered textbox on slide

### Advanced Font and Color Management

#### Font Preservation System
- **Original Font Recording**: Captures existing font and color before insertion
- **Dual Formatting**: Applies "Segoe UI Symbol" to inserted symbol, preserves original formatting for continued typing
- **Zero-Width Spacer**: Uses invisible `\u200B` character to maintain cursor positioning and formatting context
- **Multi-Byte Character Support**: Properly handles Unicode surrogate pairs (like arrow symbols)

#### Color System
- **BGR Color Conversion**: PowerPoint COM API uses BGR format internally, not RGB
- **Helper Method**: `ConvertRgbToBgr()` ensures correct color display
- **Context-Aware Colors**: Colors work in tables, text boxes, shapes, and slide insertions

### Unicode Character Handling
- **Single-Byte Characters**: Harvey balls, basic icons (1 UTF-16 code unit)
- **Multi-Byte Characters**: Arrow symbols (2 UTF-16 code units via surrogate pairs)
- **Dynamic Length Calculation**: `element.Length` determines proper character range for formatting

## Implementation Details

### ElementHelper Architecture

```csharp
ElementHelper.InsertElement(string element, int? colorRgb = null)
```

**Core Method Features**:
- Optional color parameter for colored symbols (stoplights)
- Automatic context detection (table, text, shape, slide)
- Font preservation with zero-width spacer technique
- Multi-byte Unicode character support

**Insertion Flow**:
1. **Context Detection**: Identifies current PowerPoint selection type
2. **Original Format Capture**: Records existing font and color
3. **Symbol + Spacer Insertion**: Inserts `element + "\u200B"`
4. **Dual Formatting**: Applies symbol font to element, original font to spacer
5. **Cursor Positioning**: Places cursor in formatted spacer for continued typing

### Ribbon Integration

**Direct Character Literals**: All ribbon callbacks use direct Unicode characters for clarity:
```csharp
ElementHelper.InsertElement("⭘");      // Harvey Ball 0
ElementHelper.InsertElement("🡹", 0x007748); // Green stoplight
```

**Color Values**: Stoplights use specific hex colors:
- Red: `0xab0e04`
- Amber: `0xe2ad00` 
- Green: `0x007748`

### AlignmentHelper Architecture

```csharp
AlignmentHelper.AlignLeft()                    // Align shapes to left edge
AlignmentHelper.ResizeAndSpaceEvenly(string)   // Advanced resize and spacing
```

**Core Features**:
- Standard alignment (left, center, right, top, middle, bottom)
- Distribution (horizontal and vertical spacing)
- Sizing operations (same width, same height)
- Advanced resize-and-space algorithms with preservation options
- Primary alignment (align all to first selected shape)
- Position swapping for two selected objects

**Spacing Algorithms**:
1. **Even spacing**: All shapes get equal size and spacing
2. **Preserve first**: First shape size maintained, others adjusted
3. **Preserve last**: Last shape size maintained, others adjusted
4. **User input**: Interactive spacing dialog using Microsoft.VisualBasic.InputBox

**Shape Processing Flow**:
1. **Selection validation**: Checks for proper shape selection
2. **Shape sorting**: Orders shapes by position (left-to-right or top-to-bottom)
3. **Dimension calculation**: Computes total available space
4. **Proportional adjustment**: Resizes shapes according to chosen algorithm
5. **Positioning**: Places shapes with specified spacing

### KeyboardHookManager Architecture

```csharp
KeyboardHookManager.InstallHook()    // Install keyboard hook on startup
KeyboardHookManager.RemoveHook()     // Clean up on shutdown
```

**Core Features**:
- Local thread keyboard hooks (WH_KEYBOARD) for Office environment compatibility
- Ctrl+Shift+1-6 shortcut combinations mapped to alignment functions
- Key repeat filtering to prevent multiple executions
- Deprecated AppDomain.GetCurrentThreadId() used for Windows API compatibility

**Keyboard Shortcuts**:
- **Ctrl+Shift+1**: Align Left
- **Ctrl+Shift+2**: Align Center  
- **Ctrl+Shift+3**: Align Right
- **Ctrl+Shift+4**: Align Top
- **Ctrl+Shift+5**: Align Middle
- **Ctrl+Shift+6**: Align Bottom

**Hook Implementation**:
1. **Hook Installation**: Uses SetWindowsHookEx with local thread ID
2. **Message Filtering**: Processes WH_KEYBOARD messages in Office context
3. **Key Combination Detection**: Monitors Ctrl+Shift+Number patterns
4. **Function Execution**: Direct calls to AlignmentHelper methods
5. **Resource Cleanup**: Proper unhooking on add-in shutdown

### Office Interop Usage
- Uses `Marshal.GetActiveObject("PowerPoint.Application")` to access running PowerPoint instance
- Interacts with `Application.ActiveWindow.Selection` for context-aware insertions
- Manages `TextRange.Characters()` for precise formatting control
- Utilizes `TextRange.Select()` to force formatting updates

## Technical Considerations

### PowerPoint COM API Quirks
- **Color Format**: Uses BGR instead of RGB (`ConvertRgbToBgr()` required)
- **Character Ranges**: 1-indexed, length-based (`Characters(start, length)`)
- **Selection Updates**: May require `.Select()` calls to apply formatting changes
- **Surrogate Pairs**: Multi-byte Unicode requires proper length calculation

### Font Management
- **Primary Font**: "Segoe UI Symbol" for all inserted symbols
- **Fallback Strategy**: Preserves user's original font for continued typing
- **Context Sensitivity**: Different handling for different insertion contexts (table vs text vs slide)

### Keyboard Hook Implementation
- **Hook Type**: Local thread (WH_KEYBOARD) instead of low-level global hook
- **Thread Compatibility**: Uses deprecated AppDomain.GetCurrentThreadId() for Win32 API compatibility
- **Office Integration**: Works reliably within PowerPoint's COM environment
- **Key Filtering**: Prevents key repeat events and processes only initial key presses

## Deployment & Distribution

### MSI Installer Package
The project uses MSI installer with Visual Studio Setup Project:

- **Distribution**: MSI packages via GitHub Releases
- **Installation**: Windows Installer technology with registry entries
- **VSTO Registration**: Automatic registry setup for add-in discovery
- **Prerequisites**: .NET Framework 4.7.2, VSTO Runtime, PowerPoint 2016+

### GitHub Actions Workflow
Automated build and release process (`.github/workflows/release.yml`):

1. **Project Build**: Compiles solution and setup project in Release configuration
2. **MSI Generation**: Visual Studio Setup Project creates MSI installer
3. **Version Management**: Extracts version from tags or uses date-based versioning
4. **GitHub Release**: Publishes MSI package to GitHub Releases with installation instructions

### Manual Release Trigger
- **Workflow**: Triggered manually via GitHub Actions "workflow_dispatch" or on version tags
- **Tag-based releases**: Push tags like `v1.0.0` for versioned releases  
- **Manual releases**: Use workflow_dispatch for development/testing releases

### Installer Components
- **decksterity.msi**: Windows Installer package with registry entries
- **Setup Project**: Visual Studio installer project with VSTO registration
- **Registry Entries**: HKLM\SOFTWARE\Microsoft\Office\PowerPoint\Addins\decksterity
- **VSTO Files**: Main add-in DLL, manifest, and dependencies

## Current Status

- ✅ **Harvey Balls**: Fully implemented (0-4 levels)
- ✅ **Stoplights**: Fully implemented with correct colors
- ✅ **Icons**: Fully implemented (check, cross, plus, minus, question, ellipsis)  
- ✅ **Arrows**: Fully implemented (8 directions with multi-byte support)
- ✅ **Layout Tools**: Fully implemented with comprehensive alignment and spacing features
- ✅ **Font Preservation**: Advanced system for maintaining user formatting
- ✅ **Color Support**: Full color support across all contexts
- ✅ **MSI Installer**: Professional Windows Installer with VSTO registry setup
- ✅ **Keyboard Shortcuts**: Ctrl+Shift+1-6 shortcuts for rapid alignment operations
- ✅ **Public Distribution**: MSI distribution via GitHub Releases with automated packaging