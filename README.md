# Decksterity

Decksterity is a free and open source PowerPoint add-in that aims to replicate core consulting toolbar elements, 
inspired by add-ins like [Instrumenta](https://github.com/iappyx/Instrumenta) or [Design Toolbars](https://designtoolbars.com). While Decksterity is more limited than both of these options, it includes the elements and alignment tools that I most frequently leverage conveniently centralized to one ribbon tab with familiar, easily learned hotkeys. 

Decksterity works best when combined with [**think-cell**](https://www.think-cell.com/en) to provide key missing consulting toolbox functionality - e.g., charts and graphs, agenda slides, chevrons, save/send slides, and much more. Additionally, Decksterity will be more useful if you invest in creating a [PowerPoint template](https://support.microsoft.com/en-us/office/create-and-save-a-powerpoint-template-ee4429ad-2a74-4100-82f7-50f8169c8aca) and [think-cell style file](https://www.think-cell.com/en/resources/manual/customization) suited to your needs. 

![Decksterity ribbon](assets/ribbon.png "Decksterity Ribbon")
![Decksterity access keys](assets/ribbon-with-keys.png "Ribbon hotkeys")

## Install

**[Download the latest release](https://github.com/khc32968-ui/decksterity/releases/latest)**

1. **Close PowerPoint**, then download `DecksteritySetup-<version>.msi` from the latest release
2. **Run** it. It installs for your Windows user only, so it doesn't need administrator rights. Because the installer isn't signed by a known publisher, Windows may show "Windows protected your PC": choose **More info → Run anyway**
3. **Start PowerPoint**. The first time, Office asks whether to install the Decksterity customization from an unverified publisher: choose **Install**. The **Decksterity** and **Primary** tabs then appear on the ribbon

Requirements: Windows 10 or 11 with desktop PowerPoint 2016 or later (Microsoft 365 is fine) and the .NET Framework 4.7.2 or later (included in current Windows).

**Updating**: run the newer installer; it replaces the old version. **Uninstalling**: Settings → Apps → Installed apps → *Decksterity PowerPoint Add-in* → Uninstall.

**If the tabs don't appear**: make sure PowerPoint was fully closed while installing, then check File → Options → Add-ins → *Manage: COM Add-ins* → Go, and tick *Decksterity PowerPoint Add-in*. If Office reports that the Visual Studio Tools for Office runtime is missing, install Microsoft's free [Visual Studio 2010 Tools for Office Runtime](https://www.microsoft.com/download/details.aspx?id=105522) and start PowerPoint again.

### Building from source

Needs Visual Studio 2022 or the [VS 2022 Build Tools](https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022) with the **Office/SharePoint build tools** workload. Double-click `build.bat` (or run it from a command prompt): it builds and signs the add-in into `bin\Release` and the installer into `dist\DecksteritySetup-<version>.msi`, logging to `build.log`. The first run downloads the WiX installer SDK from NuGet. The add-in's manifests are signed with a self-signed "Decksterity (local build)" certificate that `scripts\make-signing-cert.ps1` creates in your personal certificate store if you don't have one yet.

## Features

### Visual Elements
- **Harvey Balls**: Progress indicators (0-4 fill levels: ⭘, ◔, ◑, ◕, ●)
- **Arrows**: 8-directional arrows (🡹, 🡽, 🡺, 🡾, 🡻, 🡿, 🡸, 🡼)
- **Icons**: Basic symbols (✔, ✘, ➕, ➖, ❓, …)
- **Stoplights**: Colored status indicators (red, amber, green)

### Layout Tools
- **Alignment**: Left, center, right, top, middle, bottom
- **Distribution**: Even horizontal and vertical spacing
- **Sizing**: Make shapes same width/height
- **Advanced Spacing**: Resize and space evenly with preservation options (similar to ASL)
- **Primary Alignment**: Align all objects relative to the first selected object
- **Position Swapping**: Swap positions of two selected objects
- **Arrange in Grid** (`Alt-X-GR`): Tile the selected shapes into rows and columns within the selection's width, with a gap you choose; shapes keep their heights and are placed in reading order
- **Copy / Paste Position and Size** (`Alt-X-CP`, then `Alt-X-VP`): Copy one shape's position and size and apply both, or just one, to other shapes on any slide

### Color Pane (`Alt-X-CL`)
A pane docked on the right with one-click colors:
- **Fill / Line / Text rows**: click a palette color to apply it to the selected shapes (or highlighted text, or selected table cells). Each row also has *no fill* / *no line* and a `...` button for any other color
- **Line weight slider**: drag to set the outline weight of the selected shapes in 0.25 pt steps (0 removes the line), or type an exact value; it follows the current selection
- **Palette**: a fixed set of colors taken from a presentation's theme (the first time the pane opens, it uses the open presentation's theme). Change it from **Edit palette...** in the pane or the **Palette** menu (`Alt-X-PE`): add, change, reorder or remove colors, add the selected shape's color, or re-import a theme from the open presentation or any .pptx / .potx file. The palette is saved and stays the same across presentations until you change it

### Text Tools (`Alt-X-TT`)
- **Zero text margins** (`Z`): Set text insets to zero on text shapes and table cells; run again to restore PowerPoint's defaults
- **Turn off autofit** (`A`): Stop text boxes from shrinking text or resizing to fit
- **Swap text** (`S`): Swap the text of two shapes, each keeping its own formatting
- **Merge text boxes** (`M`): Combine the selected text boxes into the top-most one, keeping character formatting and bullets


## Miscellaneous

### Architecture

Decksterity is built using C# and the Visual Studio Tools for Office (VSTO) framework. Refer to this project's CLAUDE.md for more details on architecture, as well as a good starting point for adding in features.

### License & development

Decksterity is provided with an MIT license - install & use at your own risk, no warranty or support is guaranteed. However, PRs and feedback are welcome! Please feel free to submit issues if you encounter any problems.
