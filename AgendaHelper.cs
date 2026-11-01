using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Ctrl+Shift+F: agenda slides with a vertical track, a marker on the current row and the current row highlighted
    // - No agenda slide in the presentation: inserts one after the current slide
    // - Agenda slides < rows on the current agenda: duplicates it until there is one agenda slide per row, then updates them
    // - Agenda slides = rows: copies the current agenda's rows to every agenda slide and moves the highlight and marker,
    //   so the 1st agenda slide in the deck highlights row 1, the 2nd row 2, and so on
    // Agenda slides are recognised by tags on their shapes, so they survive moving, copying and duplicating slides
    public static class AgendaHelper
    {
        // Shape tag marking the agenda's parts; the agenda table identifies an agenda slide
        private const string RoleTag = "DECKSTERITY_AGENDA";
        private const string ActiveRowTag = "DECKSTERITY_AGENDA_ACTIVE";
        private const string RoleItems = "ITEMS";
        private const string RoleTrack = "TRACK";
        private const string RoleMarker = "MARKER";
        private const string RoleTitle = "TITLE";

        private const string NoStyleNoGrid = "{2D5ABB26-0587-4C30-8999-92F81FD0307C}";

        // Layout in points on a 960 x 540 (16:9) slide; scaled to the presentation's slide size
        private const float ReferenceWidth = 960f;
        private const float ReferenceHeight = 540f;
        private const float TitleLeft = 28f, TitleTop = 77f, TitleWidth = 110f, TitleHeight = 16f;
        private const float TrackX = 127f, TrackTop = 69f, TrackBottom = 516f, TrackWeight = 1.5f;
        private const float ItemsLeft = 155f, ItemsTop = 84f, ItemsWidth = 334f, RowHeight = 60f;
        private const float MarkerWidth = 11f;
        private const float ItemFontSize = 18f;
        private const float DividerWeight = 0.75f;

        private static readonly string[] DefaultItems = { "Agenda item 1", "Agenda item 2", "Agenda item 3", "Agenda item 4" };

        private static Application App => Globals.ThisAddIn.Application;

        #region Command

        public static void InsertOrUpdateAgenda()
        {
            if (App.Presentations.Count == 0) return;

            var current = GetCurrentSlide();
            if (current == null)
            {
                ShowInfo("Go to a slide in Normal view first.");
                return;
            }

            try
            {
                var presentation = App.ActivePresentation;
                var agendaSlides = FindAgendaSlides(presentation);
                App.StartNewUndoEntry();

                if (agendaSlides.Count == 0)
                {
                    InsertAgendaSlide(current);
                    return;
                }

                // The selected agenda is the one whose rows are copied; off an agenda slide, use the first one
                var source = agendaSlides.FirstOrDefault(s => s.SlideID == current.SlideID) ?? agendaSlides[0];
                int rowCount = FindShape(source, RoleItems).Table.Rows.Count;

                if (agendaSlides.Count > rowCount)
                {
                    GoTo(source);
                    string slideNumbers = string.Join(", ", agendaSlides.Select(s => s.SlideIndex));
                    ShowInfo($"This agenda has {rowCount} row{(rowCount == 1 ? "" : "s")}, but the presentation has " +
                        $"{agendaSlides.Count} agenda slides (slides {slideNumbers}).\n\n" +
                        "Delete the extra agenda slides, or add rows to the agenda (press Tab in its last cell), " +
                        "then press Ctrl+Shift+F again.");
                    return;
                }

                if (agendaSlides.Count < rowCount)
                {
                    // Copies land right after the source; move each one before its section and press Ctrl+Shift+F again
                    for (int i = agendaSlides.Count; i < rowCount; i++)
                    {
                        source.Duplicate();
                    }
                    agendaSlides = FindAgendaSlides(presentation);
                }

                FlowAgenda(source, agendaSlides);
                GoTo(source);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Agenda command failed: {ex}");
                MessageBox.Show("Couldn't update the agenda.\n\n" + ex.Message, "Decksterity",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }

        // Recolors every agenda slide in the active presentation (used by the agenda colors dialog)
        public static int ApplyStyleToPresentation(AgendaStyle style)
        {
            if (App.Presentations.Count == 0) return 0;
            var agendaSlides = FindAgendaSlides(App.ActivePresentation);
            if (agendaSlides.Count == 0) return 0;

            App.StartNewUndoEntry();
            for (int i = 0; i < agendaSlides.Count; i++)
            {
                var slide = agendaSlides[i];
                var items = FindShape(slide, RoleItems);
                int active = GetActiveRow(items, i + 1);
                ApplyRowFormatting(items.Table, active, style);
                ColorTrackAndMarker(FindShape(slide, RoleTrack), FindShape(slide, RoleMarker), style);
            }
            return agendaSlides.Count;
        }

        public static bool PresentationHasAgenda()
        {
            try
            {
                return App.Presentations.Count > 0 && FindAgendaSlides(App.ActivePresentation).Count > 0;
            }
            catch
            {
                return false;
            }
        }

        #endregion

        #region Insert

        private static void InsertAgendaSlide(Slide current)
        {
            var presentation = App.ActivePresentation;
            var style = AgendaStyle.Load();
            float sx = presentation.PageSetup.SlideWidth / ReferenceWidth;
            float sy = presentation.PageSetup.SlideHeight / ReferenceHeight;

            int index = current.SlideIndex + 1;
            var layout = FindPlainLayout(current);
            Slide slide = layout != null
                ? presentation.Slides.AddSlide(index, layout)
                : presentation.Slides.Add(index, PpSlideLayout.ppLayoutBlank);

            // The new slide's placeholders are empty; the agenda brings its own title
            foreach (var placeholder in slide.Shapes.Cast<Shape>().Where(s => s.Type == Office.MsoShapeType.msoPlaceholder).ToList())
            {
                placeholder.Delete();
            }

            AddTitle(slide, sx, sy, style);
            var track = AddTrack(slide, TrackX * sx, TrackTop * sy, (TrackBottom - TrackTop) * sy, TrackWeight);

            int rowCount = DefaultItems.Length;
            float rowHeight = RowHeight * sy;
            var items = slide.Shapes.AddTable(rowCount, 1, ItemsLeft * sx, ItemsTop * sy, ItemsWidth * sx, rowCount * rowHeight);
            items.Name = "Agenda Items";
            Tag(items, RoleItems);

            var table = items.Table;
            table.ApplyStyle(NoStyleNoGrid, false);
            table.FirstRow = false;
            table.HorizBanding = false;
            foreach (Column column in table.Columns)
            {
                column.Width = ItemsWidth * sx;
            }

            for (int row = 1; row <= rowCount; row++)
            {
                var cell = table.Cell(row, 1);
                cell.Shape.Fill.Visible = Office.MsoTriState.msoFalse;

                var frame = cell.Shape.TextFrame;
                frame.MarginLeft = 10f * sx;
                frame.MarginRight = 4f;
                frame.MarginTop = 2f;
                frame.MarginBottom = 2f;
                frame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                frame.WordWrap = Office.MsoTriState.msoTrue;

                var text = frame.TextRange;
                text.Text = DefaultItems[row - 1];
                text.Font.Size = ItemFontSize;
                text.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                GetRow(table, row).Height = rowHeight;
            }

            var marker = AddMarker(slide, track.Left - TrackWeight / 2, MarkerWidth * sx);

            ApplyRowFormatting(table, 1, style);
            ColorTrackAndMarker(track, marker, style);
            PositionMarker(marker, items, 1);
            SetActiveRow(items, 1);

            GoTo(slide);
            Select(table.Cell(1, 1).Shape.TextFrame.TextRange);
        }

        // A "Blank" layout, then "Title Only", from the current slide's master; otherwise the current slide's own layout
        private static CustomLayout FindPlainLayout(Slide current)
        {
            try
            {
                var layouts = current.CustomLayout.Design.SlideMaster.CustomLayouts.Cast<CustomLayout>().ToList();
                return layouts.FirstOrDefault(l => l.Name.IndexOf("blank", StringComparison.OrdinalIgnoreCase) >= 0)
                    ?? layouts.FirstOrDefault(l => l.Name.IndexOf("title only", StringComparison.OrdinalIgnoreCase) >= 0)
                    ?? current.CustomLayout;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not pick a layout for the agenda slide: {ex.Message}");
                return null;
            }
        }

        private static Shape AddTitle(Slide slide, float sx, float sy, AgendaStyle style)
        {
            var title = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                TitleLeft * sx, TitleTop * sy, TitleWidth * sx, TitleHeight * sy);
            title.Name = "Agenda Title";
            Tag(title, RoleTitle);

            var frame = title.TextFrame;
            frame.WordWrap = Office.MsoTriState.msoFalse;
            frame.AutoSize = PpAutoSize.ppAutoSizeNone;
            frame.MarginLeft = frame.MarginRight = frame.MarginTop = frame.MarginBottom = 0;

            var text = frame.TextRange;
            text.Text = "AGENDA";
            text.Font.Size = 10;
            text.Font.Bold = Office.MsoTriState.msoTrue;
            text.Font.Color.RGB = ElementHelper.ConvertRgbToBgr(style.TextColor);
            title.TextFrame2.TextRange.Font.Spacing = 3f; // letter-spaced, as in "A G E N D A"
            return title;
        }

        private static Shape AddTrack(Slide slide, float x, float top, float height, float weight)
        {
            var track = slide.Shapes.AddLine(x, top, x, top + height);
            track.Name = "Agenda Track";
            track.Line.Weight = weight;
            Tag(track, RoleTrack);
            return track;
        }

        private static Shape AddMarker(Slide slide, float left, float width)
        {
            var marker = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, 0, width, 10);
            marker.Name = "Agenda Marker";
            marker.Line.Visible = Office.MsoTriState.msoFalse;
            Tag(marker, RoleMarker);
            return marker;
        }

        #endregion

        #region Flow

        // Makes every agenda slide match the source's rows, columns, sizes and colors,
        // and highlights row k on the k-th agenda slide in the deck
        private static void FlowAgenda(Slide source, List<Slide> agendaSlides)
        {
            var sourceItems = FindShape(source, RoleItems);
            var sourceTable = sourceItems.Table;
            var style = AgendaStyle.FromAgenda(sourceItems, FindShape(source, RoleTrack), FindShape(source, RoleMarker));

            int rowCount = sourceTable.Rows.Count;
            int columnCount = sourceTable.Columns.Count;
            var rowHeights = sourceTable.Rows.Cast<Row>().Select(r => r.Height).ToArray();
            var columnWidths = sourceTable.Columns.Cast<Column>().Select(c => c.Width).ToArray();
            var texts = new string[rowCount, columnCount];
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    texts[row - 1, column - 1] = sourceTable.Cell(row, column).Shape.TextFrame.TextRange.Text;
                }
            }

            var sourceFont = sourceTable.Cell(1, 1).Shape.TextFrame.TextRange.Font;
            float fontSize = sourceFont.Size;
            string fontName = sourceFont.Name;

            var sourceTrack = FindShape(source, RoleTrack);
            var sourceMarker = FindShape(source, RoleMarker);

            for (int k = 0; k < agendaSlides.Count; k++)
            {
                var slide = agendaSlides[k];
                int active = Math.Min(k + 1, rowCount);
                bool isSource = slide.SlideID == source.SlideID;

                var items = isSource ? sourceItems : FindShape(slide, RoleItems);
                var table = items.Table;
                var track = FindShape(slide, RoleTrack);
                var marker = FindShape(slide, RoleMarker);

                if (!isSource)
                {
                    SetGridSize(table, rowCount, columnCount);
                    for (int column = 1; column <= columnCount; column++)
                    {
                        GetColumn(table, column).Width = columnWidths[column - 1];
                    }
                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int column = 1; column <= columnCount; column++)
                        {
                            var text = table.Cell(row, column).Shape.TextFrame.TextRange;
                            text.Text = texts[row - 1, column - 1];
                            text.Font.Size = fontSize;
                            if (!string.IsNullOrEmpty(fontName) && text.Font.Name != fontName) text.Font.Name = fontName;
                        }
                    }
                    for (int row = 1; row <= rowCount; row++)
                    {
                        GetRow(table, row).Height = rowHeights[row - 1];
                    }
                    items.Left = sourceItems.Left;
                    items.Top = sourceItems.Top;

                    track = MatchShape(slide, track, sourceTrack,
                        () => AddTrack(slide, sourceTrack.Left, sourceTrack.Top, sourceTrack.Height, sourceTrack.Line.Weight));
                    if (track != null) track.Line.Weight = sourceTrack.Line.Weight;

                    marker = MatchShape(slide, marker, sourceMarker,
                        () => AddMarker(slide, sourceMarker.Left, sourceMarker.Width));
                }

                ApplyRowFormatting(table, active, style);
                ColorTrackAndMarker(track, marker, style);
                if (marker != null) PositionMarker(marker, items, active);
                SetActiveRow(items, active);
            }
        }

        // Mirrors the source's track / marker on a target slide: creates it if missing, removes it if the source has none
        private static Shape MatchShape(Slide slide, Shape target, Shape source, Func<Shape> create)
        {
            if (source == null)
            {
                target?.Delete();
                return null;
            }

            if (target == null) target = create();
            target.Left = source.Left;
            target.Top = source.Top;
            target.Width = source.Width;
            target.Height = source.Height;
            return target;
        }

        private static void SetGridSize(Table table, int rowCount, int columnCount)
        {
            while (table.Rows.Count < rowCount) table.Rows.Add();
            while (table.Rows.Count > rowCount) GetRow(table, table.Rows.Count).Delete();
            while (table.Columns.Count < columnCount) table.Columns.Add();
            while (table.Columns.Count > columnCount) GetColumn(table, table.Columns.Count).Delete();
        }

        #endregion

        #region Formatting

        // Active row: bold in the highlight color; other rows: regular text color; a divider under every row but the last
        private static void ApplyRowFormatting(Table table, int activeRow, AgendaStyle style)
        {
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;
            for (int row = 1; row <= rowCount; row++)
            {
                bool active = row == activeRow;
                for (int column = 1; column <= columnCount; column++)
                {
                    var cell = table.Cell(row, column);
                    var font = cell.Shape.TextFrame.TextRange.Font;
                    font.Bold = active ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    font.Color.RGB = ElementHelper.ConvertRgbToBgr(active ? style.ActiveTextColor : style.TextColor);

                    var border = cell.Borders[PpBorderType.ppBorderBottom];
                    if (row < rowCount)
                    {
                        border.Visible = Office.MsoTriState.msoTrue;
                        border.Transparency = 0f;
                        border.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(style.DividerColor);
                        border.Weight = DividerWeight;
                    }
                    else
                    {
                        HideBorder(border);
                    }
                }
            }
        }

        // Setting Visible to false doesn't reliably hide a table border, so it is also made fully transparent
        private static void HideBorder(LineFormat border)
        {
            try
            {
                border.Transparency = 1f;
                border.Visible = Office.MsoTriState.msoFalse;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not hide table border: {ex.Message}");
            }
        }

        private static void ColorTrackAndMarker(Shape track, Shape marker, AgendaStyle style)
        {
            if (track != null)
            {
                track.Line.Visible = Office.MsoTriState.msoTrue;
                track.Line.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(style.TrackColor);
            }
            if (marker != null)
            {
                marker.Fill.Visible = Office.MsoTriState.msoTrue;
                marker.Fill.Solid();
                marker.Fill.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(style.MarkerColor);
            }
        }

        // Covers the active row's height, next to the track
        private static void PositionMarker(Shape marker, Shape items, int activeRow)
        {
            float top = items.Top;
            int row = 1;
            foreach (Row r in items.Table.Rows)
            {
                if (row == activeRow)
                {
                    marker.Top = top;
                    marker.Height = r.Height;
                    return;
                }
                top += r.Height;
                row++;
            }
        }

        #endregion

        #region Agenda slides and tags

        // Agenda slides in deck order: slides with a table tagged as agenda items
        private static List<Slide> FindAgendaSlides(Presentation presentation)
        {
            var slides = new List<Slide>();
            foreach (Slide slide in presentation.Slides)
            {
                if (FindShape(slide, RoleItems) != null) slides.Add(slide);
            }
            return slides;
        }

        private static Shape FindShape(Slide slide, string role)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (GetTag(shape.Tags, RoleTag) == role)
                {
                    if (role == RoleItems && shape.HasTable != Office.MsoTriState.msoTrue) continue;
                    return shape;
                }
            }
            return null;
        }

        private static void Tag(Shape shape, string role) => shape.Tags.Add(RoleTag, role);

        private static void SetActiveRow(Shape items, int row) =>
            items.Tags.Add(ActiveRowTag, row.ToString(CultureInfo.InvariantCulture));

        // The row this slide highlights, as recorded when it was last updated
        internal static int GetActiveRow(Shape items, int fallback)
        {
            int rowCount = items.Table.Rows.Count;
            if (int.TryParse(GetTag(items.Tags, ActiveRowTag), NumberStyles.Integer, CultureInfo.InvariantCulture, out int row)
                && row >= 1 && row <= rowCount)
            {
                return row;
            }
            return Math.Max(1, Math.Min(fallback, rowCount));
        }

        // Reads a tag by walking the collection (tag names are stored upper-case); null if missing
        private static string GetTag(Tags tags, string name)
        {
            for (int i = 1; i <= tags.Count; i++)
            {
                if (string.Equals(tags.Name(i), name, StringComparison.OrdinalIgnoreCase)) return tags.Value(i);
            }
            return null;
        }

        #endregion

        #region Helpers

        private static Row GetRow(Table table, int index)
        {
            int i = 1;
            foreach (Row row in table.Rows)
            {
                if (i++ == index) return row;
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        private static Column GetColumn(Table table, int index)
        {
            int i = 1;
            foreach (Column column in table.Columns)
            {
                if (i++ == index) return column;
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        // The slide in the editing pane, or the first slide selected in the thumbnail pane
        private static Slide GetCurrentSlide()
        {
            try
            {
                var window = App.ActiveWindow;
                var selection = window.Selection;
                if (selection.Type == PpSelectionType.ppSelectionSlides && selection.SlideRange.Count > 0)
                {
                    return selection.SlideRange[1];
                }
                return (Slide)window.View.Slide;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"No current slide: {ex.Message}");
                return null;
            }
        }

        private static void GoTo(Slide slide)
        {
            try
            {
                App.ActiveWindow.View.GotoSlide(slide.SlideIndex);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not go to slide: {ex.Message}");
            }
        }

        private static void Select(TextRange range)
        {
            try
            {
                range.Select();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not select agenda text: {ex.Message}");
            }
        }

        private static void ShowInfo(string message)
        {
            MessageBox.Show(message, "Decksterity", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        #endregion
    }

    // Agenda colors as 0xRRGGBB. New agendas use the saved defaults (HKCU\Software\Decksterity\Agenda);
    // updating an agenda reuses the colors already on the agenda slide it runs from
    public class AgendaStyle
    {
        private const string RegistryKeyPath = @"Software\Decksterity\Agenda";

        public const int Marker = 0, Track = 1, ActiveText = 2, Text = 3, Divider = 4;

        public static readonly string[] Labels =
        {
            "Marker (block beside the current item)",
            "Vertical line",
            "Current item text",
            "Other items text",
            "Dividers between items"
        };

        private static readonly string[] ValueNames = { "MarkerColor", "TrackColor", "ActiveTextColor", "TextColor", "DividerColor" };

        // Red #CB2026 for the marker, line and current item; black text and light gray dividers
        public static readonly int[] Defaults = { 0xCB2026, 0xCB2026, 0xCB2026, 0x000000, 0xD9D9D9 };

        public readonly int[] Colors;

        public AgendaStyle(IEnumerable<int> colors)
        {
            Colors = colors.ToArray();
        }

        public int MarkerColor => Colors[Marker];
        public int TrackColor => Colors[Track];
        public int ActiveTextColor => Colors[ActiveText];
        public int TextColor => Colors[Text];
        public int DividerColor => Colors[Divider];

        public static AgendaStyle Load()
        {
            var colors = Defaults.ToArray();
            using (var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath))
            {
                if (key != null)
                {
                    for (int i = 0; i < colors.Length; i++)
                    {
                        if (key.GetValue(ValueNames[i]) is string hex
                            && int.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int color))
                        {
                            colors[i] = color;
                        }
                    }
                }
            }
            return new AgendaStyle(colors);
        }

        public void Save()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath))
            {
                for (int i = 0; i < Colors.Length; i++)
                {
                    key.SetValue(ValueNames[i], Colors[i].ToString("X6"));
                }
            }
        }

        // Colors as they are on an existing agenda; anything that can't be read falls back to the saved defaults
        internal static AgendaStyle FromAgenda(Shape items, Shape track, Shape marker)
        {
            var style = Load();
            var colors = style.Colors;

            TryRead(() =>
            {
                if (marker != null && marker.Fill.Visible == Office.MsoTriState.msoTrue)
                    colors[Marker] = ElementHelper.ConvertRgbToBgr(marker.Fill.ForeColor.RGB);
            });
            TryRead(() =>
            {
                if (track != null && track.Line.Visible == Office.MsoTriState.msoTrue)
                    colors[Track] = ElementHelper.ConvertRgbToBgr(track.Line.ForeColor.RGB);
            });

            var table = items.Table;
            int rowCount = table.Rows.Count;
            int active = AgendaHelper.GetActiveRow(items, 1);
            TryRead(() =>
            {
                colors[ActiveText] = ElementHelper.ConvertRgbToBgr(table.Cell(active, 1).Shape.TextFrame.TextRange.Font.Color.RGB);
            });
            if (rowCount > 1)
            {
                int other = active == 1 ? 2 : 1;
                TryRead(() =>
                {
                    colors[Text] = ElementHelper.ConvertRgbToBgr(table.Cell(other, 1).Shape.TextFrame.TextRange.Font.Color.RGB);
                });
                TryRead(() =>
                {
                    var border = table.Cell(1, 1).Borders[PpBorderType.ppBorderBottom];
                    if (border.Transparency < 1f) colors[Divider] = ElementHelper.ConvertRgbToBgr(border.ForeColor.RGB);
                });
            }
            return style;
        }

        private static void TryRead(Action read)
        {
            try
            {
                read();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not read agenda color: {ex.Message}");
            }
        }
    }
}
