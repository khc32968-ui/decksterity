using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace decksterity
{
    // Centralized utility class for inserting Unicode symbols into PowerPoint slides
    // Handles context-aware insertion with font preservation and color support
    public static class ElementHelper
    {
        // Main insertion method that handles all PowerPoint selection contexts
        // Supports optional RGB color for colored symbols (stoplights)
        public static void InsertElement(string element, int? colorRgb = null)
        {
            var app = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("PowerPoint.Application");
            var selection = app.ActiveWindow.Selection;

            // Context 1: Table cell selection - insert into table with formatting preservation
            if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                Shape tableShape = null;
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        tableShape = shape;
                        break;
                    }
                }
                if (tableShape != null)
                {
                    InsertElementIntoTable(element, tableShape, selection, colorRgb);
                    return;
                }
            }

            // Context 2: Text selection - insert at cursor position with font preservation
            if (selection.Type == PpSelectionType.ppSelectionText)
            {
                InsertElementIntoText(element, selection, colorRgb);
                return;
            }
            else if (selection.Type == PpSelectionType.ppSelectionShapes)
            {
                // Context 3: Shape selection - insert into shape text frame
                foreach (Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        InsertElementIntoShapeText(element, shape, colorRgb);
                        return;
                    }
                }
            }

            // Context 4: Fallback - create new centered textbox on slide
            InsertElementIntoSlide(element, app, colorRgb);
        }


        #region Context-Specific Insertion Methods

        // Inserts element into table cell with dual font formatting system
        // Preserves original font for continued typing via zero-width spacer
        private static void InsertElementIntoTable(string element, Shape tableShape, Selection selection, int? colorRgb = null)
        {
            // Active text cursor in table cell - use advanced font preservation
            if (selection.Type == PpSelectionType.ppSelectionText && selection.TextRange != null)
            {
                // Step 1: Record original formatting before making changes
                var originalFontName = selection.TextRange.Font.Name;
                var originalFontColor = selection.TextRange.Font.Color.RGB;
                
                // Step 2: Insert symbol + zero-width space for cursor positioning
                var textRange = selection.TextRange;
                textRange.Text = element + "\u200B"; // Zero-width space
                
                // Step 3: Calculate the length of the element (handles surrogate pairs)
                int elementLength = element.Length;
                
                // Step 4: Apply symbol formatting to the element part only
                var symbolRange = textRange.Characters(1, elementLength);
                symbolRange.Font.Name = "Segoe UI Symbol";
                if (colorRgb.HasValue)
                {
                    int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                    symbolRange.Font.Color.RGB = bgrColor;
                }
                
                // Step 5: Apply original formatting to the zero-width space for future typing
                var spacerRange = textRange.Characters(elementLength + 1, 1);
                spacerRange.Font.Name = originalFontName;
                spacerRange.Font.Color.RGB = originalFontColor;
                
                // Step 6: Position cursor after the symbol but in the formatted spacer
                spacerRange.Select();
            }
            else
            {
                // No active cursor - fallback to first cell replacement
                var cell = tableShape.Table.Cell(1, 1);
                var textRange = cell.Shape.TextFrame.TextRange;
                
                // Simple replacement without cursor positioning
                textRange.Text = element;
                textRange.Font.Name = "Segoe UI Symbol";
                if (colorRgb.HasValue)
                {
                    int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                    textRange.Font.Color.RGB = bgrColor;
                }
            }
        }

        // Inserts element into active text selection with font preservation
        // Uses zero-width spacer technique for seamless cursor positioning
        private static void InsertElementIntoText(string element, Selection selection, int? colorRgb = null)
        {
            // Record original formatting before making changes
            var originalFontName = selection.TextRange.Font.Name;
            var originalFontColor = selection.TextRange.Font.Color.RGB;
            
            // Insert symbol + zero-width space for cursor positioning
            var textRange = selection.TextRange;
            textRange.Text = element + "\u200B"; // Zero-width space
            
            // Calculate the length of the element (handles surrogate pairs)
            int elementLength = element.Length;
            
            // Apply symbol formatting to the element part only
            var symbolRange = textRange.Characters(1, elementLength);
            symbolRange.Font.Name = "Segoe UI Symbol";
            if (colorRgb.HasValue)
            {
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                symbolRange.Font.Color.RGB = bgrColor;
            }
            
            // Apply original formatting to the zero-width space for future typing
            var spacerRange = textRange.Characters(elementLength + 1, 1);
            spacerRange.Font.Name = originalFontName;
            spacerRange.Font.Color.RGB = originalFontColor;
            
            // Position cursor after the symbol but in the formatted spacer
            spacerRange.Select();
        }

        // Inserts element into shape text frame replacing existing content
        // Uses dual formatting but without cursor positioning (shape context)
        private static void InsertElementIntoShapeText(string element, Shape shape, int? colorRgb = null)
        {
            // Record original formatting before making changes
            var textRange = shape.TextFrame.TextRange;
            var originalFontName = textRange.Font.Name;
            var originalFontColor = textRange.Font.Color.RGB;
            
            // Insert symbol + zero-width space for cursor positioning
            textRange.Text = element + "\u200B"; // Zero-width space
            
            // Calculate the length of the element (handles surrogate pairs)
            int elementLength = element.Length;
            
            // Apply symbol formatting to the element part only
            var symbolRange = textRange.Characters(1, elementLength);
            symbolRange.Font.Name = "Segoe UI Symbol";
            if (colorRgb.HasValue)
            {
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                symbolRange.Font.Color.RGB = bgrColor;
            }
            
            // Apply original formatting to the zero-width space for future typing
            var spacerRange = textRange.Characters(elementLength + 1, 1);
            spacerRange.Font.Name = originalFontName;
            spacerRange.Font.Color.RGB = originalFontColor;
        }

        // Creates new centered textbox on slide with symbol
        // Fallback method when no suitable selection context is found
        private static void InsertElementIntoSlide(string element, Application app, int? colorRgb = null)
        {
            // Insert as new shape in the center of the active slide
            var slide = app.ActiveWindow.View.Slide;
            float left = app.ActivePresentation.PageSetup.SlideWidth / 2 - 20;
            float top = app.ActivePresentation.PageSetup.SlideHeight / 2 - 20;
            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, 40, 40);
            
            // Configure text properties
            shape.TextFrame.TextRange.Text = element;
            shape.TextFrame.TextRange.Font.Name = "Segoe UI Symbol";
            shape.TextFrame.TextRange.Font.Size = 16;
            
            // Apply color if specified
            if (colorRgb.HasValue)
            {
                // Convert RGB to BGR format for PowerPoint
                int bgrColor = ConvertRgbToBgr(colorRgb.Value);
                shape.TextFrame.TextRange.Font.Color.RGB = bgrColor;
            }
            
            // Center and middle align the text
            shape.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
            shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
        }

        // Converts RGB color format to BGR format required by PowerPoint COM API
        // PowerPoint internally uses BGR instead of standard RGB format
        internal static int ConvertRgbToBgr(int rgbColor)
        {
            // Step 1: Extract RGB components from input color
            int r = (rgbColor >> 16) & 0xFF;
            int g = (rgbColor >> 8) & 0xFF;
            int b = rgbColor & 0xFF;
            
            // Step 2: Recombine as BGR for PowerPoint compatibility
            return (b << 16) | (g << 8) | r;
        }

        #endregion
    }
}