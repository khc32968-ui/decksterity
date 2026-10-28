using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // The color pane's palette: a fixed list of colors, usually imported once from a presentation's theme
    // Colors are stored as 0xRRGGBB in HKCU\Software\Decksterity\Palette and survive restarts
    public static class PaletteStore
    {
        private const string RegistryKeyPath = @"Software\Decksterity\Palette";
        private const string ColorsValue = "Colors";
        private const string SourceValue = "Source";

        // Office 2013+ default theme, used only if nothing has been saved and no presentation is open
        private static readonly int[] OfficeDefault =
        {
            0xFFFFFF, 0x000000, 0xE7E6E6, 0x44546A,
            0x4472C4, 0xED7D31, 0xA5A5A5, 0xFFC000, 0x5B9BD5, 0x70AD47
        };

        // Theme slots in the order PowerPoint's own color picker shows them
        private static readonly Office.MsoThemeColorSchemeIndex[] ThemeSlots =
        {
            Office.MsoThemeColorSchemeIndex.msoThemeLight1,
            Office.MsoThemeColorSchemeIndex.msoThemeDark1,
            Office.MsoThemeColorSchemeIndex.msoThemeLight2,
            Office.MsoThemeColorSchemeIndex.msoThemeDark2,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent1,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent2,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent3,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent4,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent5,
            Office.MsoThemeColorSchemeIndex.msoThemeAccent6
        };

        private static List<int> colors;
        private static string source;

        // Raised after the palette is saved so open panes can redraw
        public static event EventHandler Changed;

        // Current palette as 0xRRGGBB values
        public static IReadOnlyList<int> Colors
        {
            get
            {
                EnsureLoaded();
                return colors;
            }
        }

        // Where the palette came from, e.g. "Theme of Template.potx"
        public static string Source
        {
            get
            {
                EnsureLoaded();
                return source;
            }
        }

        public static void Save(IEnumerable<int> newColors, string newSource)
        {
            colors = newColors.ToList();
            source = newSource ?? "";
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                key.SetValue(ColorsValue, string.Join(",", colors.Select(c => c.ToString("X6"))));
                key.SetValue(SourceValue, source);
            }
            Changed?.Invoke(null, EventArgs.Empty);
        }

        // Reads the 10 theme colors (backgrounds, text, accents) of a presentation's first slide master
        public static List<int> ReadThemeColors(Presentation presentation)
        {
            var scheme = presentation.SlideMaster.Theme.ThemeColorScheme;
            return ThemeSlots.Select(slot => ElementHelper.ConvertRgbToBgr(scheme.Colors(slot).RGB)).ToList();
        }

        // Opens a .pptx/.potx file without a window, reads its theme colors, and closes it
        public static List<int> ReadThemeColors(string path)
        {
            var presentation = Globals.ThisAddIn.Application.Presentations.Open(path,
                ReadOnly: Office.MsoTriState.msoTrue, Untitled: Office.MsoTriState.msoFalse, WithWindow: Office.MsoTriState.msoFalse);
            try
            {
                return ReadThemeColors(presentation);
            }
            finally
            {
                presentation.Close();
            }
        }

        public static IReadOnlyList<int> DefaultColors => OfficeDefault;

        public static string ToHex(int rgb) => "#" + rgb.ToString("X6");

        public static System.Drawing.Color ToDrawingColor(int rgb) =>
            System.Drawing.Color.FromArgb((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);

        public static int FromDrawingColor(System.Drawing.Color color) => (color.R << 16) | (color.G << 8) | color.B;

        // First use: take the open presentation's theme (so the palette matches the deck you're in) and save it
        private static void EnsureLoaded()
        {
            if (colors != null) return;

            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                var saved = key?.GetValue(ColorsValue) as string;
                if (saved != null)
                {
                    colors = saved.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(hex => int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int c) ? c : -1)
                        .Where(c => c >= 0)
                        .ToList();
                    source = key.GetValue(SourceValue) as string ?? "";
                    return;
                }
            }

            try
            {
                var presentation = Globals.ThisAddIn.Application.ActivePresentation;
                Save(ReadThemeColors(presentation), "Theme of " + presentation.Name);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not read the active presentation's theme: {ex.Message}");
                Save(OfficeDefault, "Office default theme");
            }
        }
    }
}
