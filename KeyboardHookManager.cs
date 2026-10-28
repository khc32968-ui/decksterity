using System;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace decksterity
{
    // Manages keyboard shortcuts for Decksterity PowerPoint add-in
    // Implements Windows local thread keyboard hooks for Ctrl+Shift and Ctrl+Arrow combinations
    public class KeyboardHookManager
    {
        // Windows API imports
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int nVirtKey);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        // Hook constants
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        // lParam flags for WH_KEYBOARD messages
        private const long PreviousStateFlag = 1L << 30;   // key was already down (auto-repeat)
        private const long TransitionStateFlag = 1L << 31; // key is being released

        // Delegate and hook variables
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private readonly LowLevelKeyboardProc keyboardProc = HookCallback;
        private static IntPtr keyboardHook = IntPtr.Zero;
        private static KeyboardHookManager instance;

        // Installs local thread keyboard hook for capturing Ctrl+Shift shortcuts
        public void InstallHook()
        {
            try
            {
                instance = this;
                keyboardHook = SetWindowsHookEx(WH_KEYBOARD, keyboardProc,
                    IntPtr.Zero, (uint)AppDomain.GetCurrentThreadId());

                if (keyboardHook == IntPtr.Zero)
                {
                    System.Diagnostics.Debug.WriteLine("Failed to install keyboard hook");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error installing keyboard hook: {ex.Message}");
            }
        }

        // Removes keyboard hook and cleans up resources
        public void RemoveHook()
        {
            try
            {
                if (keyboardHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(keyboardHook);
                    keyboardHook = IntPtr.Zero;
                }
                instance = null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error removing keyboard hook: {ex.Message}");
            }
        }

        // Main hook callback that processes keyboard messages from Windows
        // Handled shortcuts are swallowed so PowerPoint's native binding for the same keys doesn't also fire
        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == HC_ACTION)
            {
                Keys keyData = (Keys)wParam;
                long flags = (long)lParam;
                bool isKeyUp = (flags & TransitionStateFlag) != 0;
                bool isRepeat = (flags & PreviousStateFlag) != 0;

                bool ctrl = IsKeyDown(Keys.ControlKey);
                bool shift = IsKeyDown(Keys.ShiftKey);
                bool alt = IsKeyDown(Keys.Menu); // excludes AltGr (reported as Ctrl+Alt)

                if (!isKeyUp && ctrl && !alt && IsSlideEditorFocused())
                {
                    bool handled = false;

                    // Ctrl(+Shift)+Arrow resizes shapes; holding the key repeats the resize
                    if (IsArrowKey(keyData))
                    {
                        handled = HandleResize(keyData, shift);
                    }
                    // Other shortcuts fire once per press (filter out key repeats)
                    else if (shift && !isRepeat)
                    {
                        handled = HandleShortcut(keyData);
                    }

                    if (handled)
                    {
                        return (IntPtr)1;
                    }
                }
            }

            return CallNextHookEx(keyboardHook, nCode, wParam, lParam);
        }

        // Executes the action mapped to Ctrl+Shift+<key>; returns true if the key was consumed
        private static bool HandleShortcut(Keys keyData)
        {
            try
            {
                switch (keyData)
                {
                    // Aligning / adjusting objects
                    case Keys.D1: AlignmentHelper.AlignLeft(); return true;
                    case Keys.D2: AlignmentHelper.AlignCenter(); return true;
                    case Keys.D3: AlignmentHelper.AlignRight(); return true;
                    case Keys.D4: AlignmentHelper.AlignTop(); return true;
                    case Keys.D5: AlignmentHelper.AlignMiddle(); return true;
                    case Keys.D6: AlignmentHelper.AlignBottom(); return true;
                    case Keys.D7: AlignmentHelper.DistributeVertically(); return true;
                    case Keys.D8: AlignmentHelper.DistributeHorizontally(); return true;
                    case Keys.D9: AlignmentHelper.BringToFront(); return true;
                    case Keys.OemMinus: AlignmentHelper.SendToBack(); return true;

                    // Insert slide elements
                    case Keys.O: SlideElementHelper.InsertStickyOrSetDefault(); return true;
                    case Keys.T: SlideElementHelper.InsertTable(); return true;
                    case Keys.B: SlideElementHelper.ToggleBulletsOrInsertTextBox(); return true;
                    case Keys.H: SlideElementHelper.InsertHeader(); return true;
                    case Keys.W: SlideElementHelper.InsertDownArrow(); return true;
                    case Keys.Q: SlideElementHelper.InsertQuote(); return true;
                    case Keys.A: SlideElementHelper.InsertStatusSticker(); return true;
                    case Keys.N: SlideElementHelper.InsertFootnote(); return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling keyboard shortcut: {ex.Message}");
            }
            return false;
        }

        // Ctrl+Arrow grows/shrinks selected shapes; +Shift uses the smallest increment
        // Only consumes the key when shapes are selected so text navigation and slide reordering still work
        private static bool HandleResize(Keys keyData, bool fine)
        {
            try
            {
                float step = fine ? AlignmentHelper.FineResizeStep : AlignmentHelper.ResizeStep;
                switch (keyData)
                {
                    case Keys.Up: return AlignmentHelper.ResizeSelection(0, step);
                    case Keys.Down: return AlignmentHelper.ResizeSelection(0, -step);
                    case Keys.Right: return AlignmentHelper.ResizeSelection(step, 0);
                    case Keys.Left: return AlignmentHelper.ResizeSelection(-step, 0);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error handling resize shortcut: {ex.Message}");
            }
            return false;
        }

        private static bool IsArrowKey(Keys keyData)
        {
            return keyData == Keys.Up || keyData == Keys.Down || keyData == Keys.Left || keyData == Keys.Right;
        }

        // Ignores keystrokes aimed at the ribbon, task panes and dialog text fields
        // (the slide editor itself is a "paneClassDC" window, including while editing text in a shape)
        private static bool IsSlideEditorFocused()
        {
            IntPtr focus = GetFocus();
            if (focus == IntPtr.Zero) return true;

            var className = new StringBuilder(256);
            GetClassName(focus, className, className.Capacity);
            string name = className.ToString();

            return !(name == "NetUIHWND"
                || name.IndexOf("Edit", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("Combo", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        // Checks if specified key is currently pressed using Windows API
        private static bool IsKeyDown(Keys keys)
        {
            return (GetKeyState((int)keys) & 0x8000) == 0x8000;
        }
    }
}
