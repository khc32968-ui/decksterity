using System;
using System.Globalization;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Win32;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Inserts consulting-style slide elements (stickies, tables, quotes, footnotes...) bound to Ctrl+Shift+letter
    // Elements use simple neutral formatting; the sticky's look can be customized by selecting a shape and pressing Ctrl+Shift+O
    public static class SlideElementHelper
    {
        private const float Inch = 72f;
        private const float Margin = 0.4f * Inch;
        private const string StickyRegistryKey = @"Software\Decksterity\Sticky";
        private const string FootnoteName = "Decksterity Footnote";
        private const string NoStyleTableGrid = "{5940675A-B579-460E-94D1-54222C63F5DA}";

        private static Application App => Globals.ThisAddIn.Application;

        private static Slide ActiveSlide => (Slide)App.ActiveWindow.View.Slide;

        private static float SlideWidth => App.ActivePresentation.PageSetup.SlideWidth;

        private static float SlideHeight => App.ActivePresentation.PageSetup.SlideHeight;

        #region Elements

        // Ctrl+Shift+O: inserts a sticky note, or saves the single selected shape as the default sticky
        public static void InsertStickyOrSetDefault()
        {
            var selection = App.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionShapes && selection.ShapeRange.Count == 1
                && selection.ShapeRange[1].HasTextFrame == Office.MsoTriState.msoTrue)
            {
                SaveStickyDefault(selection.ShapeRange[1]);
                MessageBox.Show("Saved the selected shape as the default sticky.", "Decksterity",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var shape = LoadStickyDefault() ?? AddDefaultSticky();
            shape.Name = "Decksterity Sticky";
            SelectAllText(shape);
        }

        // Ctrl+Shift+T: inserts a centered 3x3 table with plain gridlines
        public static void InsertTable()
        {
            float width = SlideWidth * 0.8f;
            float height = 1.2f * Inch;
            var shape = ActiveSlide.Shapes.AddTable(3, 3, (SlideWidth - width) / 2, (SlideHeight - height) / 2, width, height);
            shape.Table.ApplyStyle(NoStyleTableGrid, false);
            shape.Select();
        }

        // Ctrl+Shift+B: toggles bullets on the selected text / text shapes, otherwise inserts a new text box
        public static void ToggleBulletsOrInsertTextBox()
        {
            var selection = App.ActiveWindow.Selection;
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                ToggleBullets(selection.TextRange2);
                return;
            }

            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                bool toggled = false;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        ToggleBullets(shape.TextFrame2.TextRange);
                        toggled = true;
                    }
                }
                if (toggled) return;
            }

            float width = 4f * Inch;
            var textBox = ActiveSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (SlideWidth - width) / 2, SlideHeight / 2 - 0.25f * Inch, width, 0.5f * Inch);
            textBox.TextFrame.TextRange.Text = "Text";
            SelectAllText(textBox);
        }

        // Ctrl+Shift+H: inserts a bold column header with a rule underneath
        public static void InsertHeader()
        {
            float width = 3f * Inch;
            float left = (SlideWidth - width) / 2;
            float top = SlideHeight / 2 - 0.3f * Inch;

            var text = ActiveSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, 0.4f * Inch);
            text.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            text.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;
            text.TextFrame.TextRange.Text = "Header";
            text.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
            text.TextFrame.TextRange.Font.Size = 14;
            text.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

            float lineTop = text.Top + text.Height;
            var rule = ActiveSlide.Shapes.AddLine(left, lineTop, left + width, lineTop);
            rule.Line.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0x000000);
            rule.Line.Weight = 1f;

            var group = ActiveSlide.Shapes.Range(new object[] { text.Name, rule.Name }).Group();
            group.Name = "Decksterity Header";
            SelectAllText(text);
        }

        // Ctrl+Shift+W: inserts a gray block down arrow
        public static void InsertDownArrow()
        {
            float size = 0.4f * Inch;
            var shape = ActiveSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeDownArrow,
                (SlideWidth - size) / 2, (SlideHeight - size) / 2, size, size);
            shape.Fill.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0x7F7F7F);
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.Select();
        }

        // Ctrl+Shift+Q: inserts a quote box with an attribution line
        public static void InsertQuote()
        {
            float width = 4f * Inch;
            float height = 1.2f * Inch;
            var shape = ActiveSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                (SlideWidth - width) / 2, (SlideHeight - height) / 2, width, height);
            shape.Fill.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0xF2F2F2);
            shape.Line.Visible = Office.MsoTriState.msoFalse;
            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

            var text = shape.TextFrame.TextRange;
            text.Text = "“Quote”\r– Source";
            text.Font.Color.RGB = ElementHelper.ConvertRgbToBgr(0x000000);

            var quote = text.Paragraphs(1, 1);
            quote.Font.Italic = Office.MsoTriState.msoTrue;
            quote.Font.Size = 14;
            quote.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

            var source = text.Paragraphs(2, 1);
            source.Font.Size = 11;
            source.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;

            Select(quote.Characters(2, 5));
        }

        // Ctrl+Shift+A: inserts a status sticker (e.g. PRELIMINARY) in the top-right corner
        public static void InsertStatusSticker()
        {
            float width = 1.6f * Inch;
            float height = 0.35f * Inch;
            var shape = ActiveSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRoundedRectangle,
                SlideWidth - Margin - width, Margin, width, height);
            shape.Name = "Decksterity Status Sticker";
            shape.Fill.Visible = Office.MsoTriState.msoFalse;
            shape.Line.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0x595959);
            shape.Line.Weight = 1f;

            var text = shape.TextFrame.TextRange;
            text.Text = "PRELIMINARY";
            text.Font.Size = 10;
            text.Font.Bold = Office.MsoTriState.msoTrue;
            text.Font.Color.RGB = ElementHelper.ConvertRgbToBgr(0x595959);
            text.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            SelectAllText(shape);
        }

        // Ctrl+Shift+N: inserts a Note/Source footnote at the bottom-left, or jumps into the existing one
        public static void InsertFootnote()
        {
            foreach (Shape existing in ActiveSlide.Shapes)
            {
                if (existing.Name == FootnoteName)
                {
                    SelectAllText(existing);
                    return;
                }
            }

            float height = 0.35f * Inch;
            var shape = ActiveSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal,
                Margin, SlideHeight - Margin - height, SlideWidth - 2 * Margin, height);
            shape.Name = FootnoteName;
            shape.TextFrame.WordWrap = Office.MsoTriState.msoTrue;
            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorBottom;

            var text = shape.TextFrame.TextRange;
            text.Text = "Note: \rSource: ";
            text.Font.Size = 9;
            text.Font.Color.RGB = ElementHelper.ConvertRgbToBgr(0x595959);
            text.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

            // Place the cursor after "Note: "
            Select(text.Paragraphs(1, 1).Characters(7, 0));
        }

        #endregion

        #region Sticky defaults

        private static Shape AddDefaultSticky()
        {
            float width = 1.8f * Inch;
            float height = 0.5f * Inch;
            var shape = ActiveSlide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle,
                (SlideWidth - width) / 2, (SlideHeight - height) / 2, width, height);
            shape.Fill.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0xFFFF99);
            shape.Line.ForeColor.RGB = ElementHelper.ConvertRgbToBgr(0x000000);
            shape.Line.Weight = 0.75f;

            var text = shape.TextFrame.TextRange;
            text.Text = "Sticky";
            text.Font.Size = 12;
            text.Font.Color.RGB = ElementHelper.ConvertRgbToBgr(0x000000);
            text.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            return shape;
        }

        // Records the selected shape's geometry, fill, line and font so new stickies match it
        private static void SaveStickyDefault(Shape source)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(StickyRegistryKey))
            {
                var text = source.TextFrame.TextRange;
                key.SetValue("AutoShapeType", (int)source.AutoShapeType);
                SetFloat(key, "Left", source.Left);
                SetFloat(key, "Top", source.Top);
                SetFloat(key, "Width", source.Width);
                SetFloat(key, "Height", source.Height);
                key.SetValue("FillVisible", (int)source.Fill.Visible);
                key.SetValue("FillColor", source.Fill.ForeColor.RGB);
                SetFloat(key, "FillTransparency", source.Fill.Transparency);
                key.SetValue("LineVisible", (int)source.Line.Visible);
                key.SetValue("LineColor", source.Line.ForeColor.RGB);
                SetFloat(key, "LineWeight", source.Line.Weight);
                key.SetValue("Text", text.Text);
                key.SetValue("FontName", text.Font.Name);
                SetFloat(key, "FontSize", text.Font.Size);
                key.SetValue("FontBold", (int)text.Font.Bold);
                key.SetValue("FontColor", text.Font.Color.RGB);
                key.SetValue("Alignment", (int)text.ParagraphFormat.Alignment);
            }
        }

        // Recreates the saved sticky on the active slide; returns null if no default has been saved
        private static Shape LoadStickyDefault()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(StickyRegistryKey))
            {
                if (key == null) return null;

                var shape = ActiveSlide.Shapes.AddShape((Office.MsoAutoShapeType)(int)key.GetValue("AutoShapeType"),
                    GetFloat(key, "Left"), GetFloat(key, "Top"), GetFloat(key, "Width"), GetFloat(key, "Height"));
                shape.Fill.Visible = (Office.MsoTriState)(int)key.GetValue("FillVisible");
                shape.Fill.ForeColor.RGB = (int)key.GetValue("FillColor");
                shape.Fill.Transparency = GetFloat(key, "FillTransparency");
                shape.Line.Visible = (Office.MsoTriState)(int)key.GetValue("LineVisible");
                shape.Line.ForeColor.RGB = (int)key.GetValue("LineColor");
                shape.Line.Weight = GetFloat(key, "LineWeight");

                var text = shape.TextFrame.TextRange;
                text.Text = (string)key.GetValue("Text");
                text.Font.Name = (string)key.GetValue("FontName");
                text.Font.Size = GetFloat(key, "FontSize");
                text.Font.Bold = (Office.MsoTriState)(int)key.GetValue("FontBold");
                text.Font.Color.RGB = (int)key.GetValue("FontColor");
                text.ParagraphFormat.Alignment = (PpParagraphAlignment)(int)key.GetValue("Alignment");
                return shape;
            }
        }

        private static void SetFloat(RegistryKey key, string name, float value)
        {
            key.SetValue(name, value.ToString(CultureInfo.InvariantCulture));
        }

        private static float GetFloat(RegistryKey key, string name)
        {
            return float.Parse((string)key.GetValue(name), CultureInfo.InvariantCulture);
        }

        #endregion

        #region Helpers

        // Toggles bullets (with a hanging indent) based on the current state of the range
        private static void ToggleBullets(Office.TextRange2 range)
        {
            var format = range.ParagraphFormat;
            bool turnOn = format.Bullet.Visible != Office.MsoTriState.msoTrue;

            format.Bullet.Visible = turnOn ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
            if (turnOn)
            {
                format.Bullet.Type = Office.MsoBulletType.msoBulletUnnumbered;
                format.LeftIndent = 0.2f * Inch;
                format.FirstLineIndent = -0.2f * Inch;
            }
            else
            {
                format.LeftIndent = 0;
                format.FirstLineIndent = 0;
            }
        }

        // Selects all of a shape's text so typing replaces the placeholder
        private static void SelectAllText(Shape shape)
        {
            Select(shape.TextFrame.TextRange);
        }

        // Text selection fails if the slide pane isn't active; the element is still inserted
        private static void Select(TextRange range)
        {
            try
            {
                range.Select();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not select inserted text: {ex.Message}");
            }
        }

        #endregion
    }
}
