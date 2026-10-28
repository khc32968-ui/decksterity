using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace decksterity
{
    // Content of the docked color pane: one-click palette swatches for fill, line and text, and a line-weight slider
    // Swatches come from PaletteStore; "Edit palette..." opens PaletteEditorForm
    public class ColorPaneControl : UserControl
    {
        private const int SwatchSize = 22;
        private const int StepsPerPoint = 4;   // slider moves in 0.25 pt steps
        private const int SliderMaxPoints = 6; // the number box accepts more

        private readonly ToolTip toolTip = new ToolTip();
        private readonly FlowLayoutPanel root;
        private readonly FlowLayoutPanel fillSwatches;
        private readonly FlowLayoutPanel lineSwatches;
        private readonly FlowLayoutPanel textSwatches;
        private readonly FlowLayoutPanel weightRow;
        private readonly TrackBar weightSlider;
        private readonly NumericUpDown weightBox;
        private readonly Label sourceLabel;
        private bool syncing; // true while the controls are being updated from the selection, so nothing is applied

        public ColorPaneControl()
        {
            BackColor = SystemColors.Window;
            Font = new Font("Segoe UI", 9f);
            AutoScaleMode = AutoScaleMode.Font;

            root = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(8, 6, 8, 8)
            };
            Controls.Add(root);

            fillSwatches = AddSection("Fill");
            lineSwatches = AddSection("Line");

            root.Controls.Add(Header("Line weight"));
            weightRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 0, 0, 6) };
            weightSlider = new TrackBar
            {
                Minimum = 0,
                Maximum = SliderMaxPoints * StepsPerPoint,
                SmallChange = 1,
                LargeChange = StepsPerPoint,
                TickStyle = TickStyle.None,
                AutoSize = false,
                Height = 26
            };
            weightBox = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 20,
                Increment = 0.25m,
                DecimalPlaces = 2,
                Width = 58
            };
            var unit = new Label { Text = "pt", AutoSize = true, Margin = new Padding(2, 6, 0, 0) };
            weightRow.Controls.AddRange(new Control[] { weightSlider, weightBox, unit });
            root.Controls.Add(weightRow);
            toolTip.SetToolTip(weightSlider, "Outline weight of the selected shapes (0 = no line)");

            textSwatches = AddSection("Text");

            var editLink = new LinkLabel { Text = "Edit palette...", AutoSize = true, Margin = new Padding(0, 10, 0, 2) };
            editLink.LinkClicked += (s, e) => PaletteEditorForm.ShowEditor();
            root.Controls.Add(editLink);

            sourceLabel = new Label { AutoSize = true, ForeColor = SystemColors.GrayText };
            root.Controls.Add(sourceLabel);

            weightSlider.MouseDown += (s, e) => Safe(FormatHelper.StartUndoEntry);
            weightSlider.KeyDown += (s, e) => Safe(FormatHelper.StartUndoEntry);
            weightSlider.ValueChanged += OnSliderChanged;
            weightBox.ValueChanged += OnWeightBoxChanged;

            root.Resize += (s, e) => FitToWidth();
            PaletteStore.Changed += OnPaletteChanged;

            BuildSwatches();
            FitToWidth();
        }

        // Updates the line-weight controls to match the current selection
        public void SyncToSelection()
        {
            float? weight = FormatHelper.GetSelectedLineWeight();
            syncing = true;
            try
            {
                weightSlider.Enabled = weightBox.Enabled = weight.HasValue;
                float value = weight ?? 0f;
                weightBox.Value = Math.Min(weightBox.Maximum, (decimal)value);
                weightSlider.Value = Math.Min(weightSlider.Maximum, (int)Math.Round(value * StepsPerPoint));
            }
            finally
            {
                syncing = false;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                PaletteStore.Changed -= OnPaletteChanged;
                toolTip.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Layout

        private FlowLayoutPanel AddSection(string title)
        {
            root.Controls.Add(Header(title));
            var swatches = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = true,
                Margin = new Padding(0, 0, 0, 6)
            };
            root.Controls.Add(swatches);
            return swatches;
        }

        private static Label Header(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Margin = new Padding(0, 6, 0, 2)
            };
        }

        // Swatch rows wrap at the pane's width; the slider takes the space left of the number box
        private void FitToWidth()
        {
            int width = Math.Max(SwatchSize * 4, root.ClientSize.Width - root.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);
            foreach (var panel in new[] { fillSwatches, lineSwatches, textSwatches })
            {
                panel.MinimumSize = new Size(width, 0);
                panel.MaximumSize = new Size(width, 0);
            }
            weightSlider.Width = Math.Max(60, width - weightBox.Width - 30);
        }

        private void OnPaletteChanged(object sender, EventArgs e)
        {
            if (IsDisposed) return;
            BuildSwatches();
        }

        private void BuildSwatches()
        {
            var palette = PaletteStore.Colors;
            PopulateSwatches(fillSwatches, palette, rgb => FormatHelper.ApplyFill(rgb), "No fill", () => FormatHelper.ApplyFill(null));
            PopulateSwatches(lineSwatches, palette, rgb => FormatHelper.ApplyLine(rgb), "No line", () => FormatHelper.ApplyLine(null));
            PopulateSwatches(textSwatches, palette, rgb => FormatHelper.ApplyTextColor(rgb), null, null);
            sourceLabel.Text = "Palette: " + PaletteStore.Source;
        }

        private void PopulateSwatches(FlowLayoutPanel panel, IReadOnlyList<int> palette, Action<int> apply, string noneLabel, Action applyNone)
        {
            panel.SuspendLayout();
            foreach (Control old in panel.Controls.Cast<Control>().ToList())
            {
                old.Dispose();
            }
            panel.Controls.Clear();

            foreach (int rgb in palette)
            {
                int color = rgb;
                var swatch = NewSwatch();
                swatch.BackColor = PaletteStore.ToDrawingColor(color);
                swatch.Click += (s, e) => Safe(() => apply(color));
                toolTip.SetToolTip(swatch, $"{PaletteStore.ToHex(color)}  (R{(color >> 16) & 0xFF} G{(color >> 8) & 0xFF} B{color & 0xFF})");
                panel.Controls.Add(swatch);
            }

            if (applyNone != null)
            {
                var none = NewSwatch();
                none.BackColor = Color.White;
                none.Paint += (s, e) =>
                {
                    using (var pen = new Pen(Color.Red, 1.5f))
                    {
                        e.Graphics.DrawLine(pen, 3, none.Height - 4, none.Width - 4, 3);
                    }
                };
                none.Click += (s, e) => Safe(applyNone);
                toolTip.SetToolTip(none, noneLabel);
                panel.Controls.Add(none);
            }

            var more = NewSwatch();
            more.Text = "...";
            more.Font = new Font("Segoe UI", 7f);
            more.Click += (s, e) =>
            {
                int? picked = PickColor(palette);
                if (picked.HasValue) Safe(() => apply(picked.Value));
            };
            toolTip.SetToolTip(more, "More colors...");
            panel.Controls.Add(more);

            panel.ResumeLayout();
        }

        private static Button NewSwatch()
        {
            var swatch = new Button
            {
                Size = new Size(SwatchSize, SwatchSize),
                Margin = new Padding(1),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                TabStop = false,
                Padding = Padding.Empty
            };
            swatch.FlatAppearance.BorderColor = Color.FromArgb(160, 160, 160);
            swatch.FlatAppearance.BorderSize = 1;
            return swatch;
        }

        // Standard color dialog with the palette preloaded as custom colors; returns 0xRRGGBB or null if cancelled
        internal static int? PickColor(IReadOnlyList<int> palette, int? initial = null)
        {
            using (var dialog = new ColorDialog { FullOpen = true, AnyColor = true })
            {
                dialog.CustomColors = palette.Take(16).Select(c => ElementHelper.ConvertRgbToBgr(c)).ToArray(); // COLORREF is BGR
                if (initial.HasValue) dialog.Color = PaletteStore.ToDrawingColor(initial.Value);
                return dialog.ShowDialog() == DialogResult.OK ? PaletteStore.FromDrawingColor(dialog.Color) : (int?)null;
            }
        }

        #endregion

        #region Line weight

        private void OnSliderChanged(object sender, EventArgs e)
        {
            if (syncing) return;
            float points = weightSlider.Value / (float)StepsPerPoint;
            syncing = true;
            weightBox.Value = (decimal)points;
            syncing = false;
            Safe(() => FormatHelper.ApplyLineWeight(points));
        }

        private void OnWeightBoxChanged(object sender, EventArgs e)
        {
            if (syncing) return;
            float points = (float)weightBox.Value;
            syncing = true;
            weightSlider.Value = Math.Min(weightSlider.Maximum, (int)Math.Round(points * StepsPerPoint));
            syncing = false;
            Safe(() =>
            {
                FormatHelper.StartUndoEntry();
                FormatHelper.ApplyLineWeight(points);
            });
        }

        #endregion

        // A click with no presentation open (or a COM hiccup) shouldn't crash PowerPoint
        private static void Safe(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Color pane action failed: {ex.Message}");
            }
        }
    }
}
