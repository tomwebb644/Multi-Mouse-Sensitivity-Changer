using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class AddDeviceForm : Form
    {
        readonly TextBox _nameTextBox;
        readonly TextBox _pathTextBox;
        readonly NumericUpDown _speedSelector;
        readonly NumericUpDown _scrollLinesSelector;
        readonly NumericUpDown _scrollCharsSelector;
        readonly NumericUpDown _doubleClickSelector;
        readonly CheckBox _enableCheckBox;
        readonly CheckBox _autoApplyCheckBox;
        readonly CheckBox _applyOnStartupCheckBox;
        readonly CheckBox _enhancePrecisionCheckBox;
        readonly CheckBox _swapButtonsCheckBox;
        readonly Label _statusLabel;
        readonly Panel _colorPreview;
        Color _selectedColor = Color.Gray;
        RawInputCaptureWindow _captureWindow;

        public Program.DeviceProfile NewProfile { get; private set; }

        public AddDeviceForm(Program.DeviceProfile existing = null, IEnumerable<Color> existingColors = null)
        {
            Text = "Add new device";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            MinimumSize = new Size(560, 0);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 6,
                Padding = new Padding(12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            for (int i = 0; i < layout.RowCount; i++)
            {
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            }

            var identityGroup = new GroupBox
            {
                Text = "Device identity",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var identityLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8)
            };
            identityLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25));
            identityLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            identityLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 20));
            identityLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            identityLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            identityLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            identityLayout.Controls.Add(new Label { Text = "Device path", AutoSize = true }, 0, 0);
            _pathTextBox = new TextBox { Dock = DockStyle.Fill };
            identityLayout.Controls.Add(_pathTextBox, 1, 0);
            var captureButton = new Button { Text = "Capture movement", AutoSize = true };
            captureButton.Click += (_, __) => BeginCapture();
            identityLayout.Controls.Add(captureButton, 2, 0);

            identityLayout.Controls.Add(new Label { Text = "Display name", AutoSize = true }, 0, 1);
            _nameTextBox = new TextBox { Dock = DockStyle.Fill };
            identityLayout.Controls.Add(_nameTextBox, 1, 1);
            identityLayout.SetColumnSpan(_nameTextBox, 2);

            identityLayout.Controls.Add(new Label { Text = "Icon color", AutoSize = true }, 0, 2);
            var colorButton = new Button { Text = "Choose color", AutoSize = true };
            colorButton.Click += (_, __) => ChooseColor();
            var colorPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoSize = true
            };
            _colorPreview = new Panel { Width = 32, Height = 32, BackColor = _selectedColor, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(8, 0, 0, 0) };
            colorPanel.Controls.Add(colorButton);
            colorPanel.Controls.Add(_colorPreview);
            identityLayout.Controls.Add(colorPanel, 1, 2);
            identityLayout.SetColumnSpan(colorPanel, 2);

            identityGroup.Controls.Add(identityLayout);
            layout.Controls.Add(identityGroup, 0, 0);

            var pointerGroup = new GroupBox
            {
                Text = "Pointer behavior",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var pointerLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8)
            };
            pointerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            pointerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 35));
            pointerLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            pointerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pointerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            pointerLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            pointerLayout.Controls.Add(new Label { Text = "Pointer speed", AutoSize = true }, 0, 0);
            _speedSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 20,
                Value = 10,
                Dock = DockStyle.Left,
                Width = 80
            };
            pointerLayout.Controls.Add(_speedSelector, 1, 0);
            var testButton = new Button { Text = "Test speed", AutoSize = true };
            testButton.Click += (_, __) => TestSpeed();
            pointerLayout.Controls.Add(testButton, 2, 0);

            _enhancePrecisionCheckBox = new CheckBox { Text = "Enhance pointer precision", AutoSize = true };
            pointerLayout.Controls.Add(_enhancePrecisionCheckBox, 0, 1);
            pointerLayout.SetColumnSpan(_enhancePrecisionCheckBox, 3);

            pointerLayout.Controls.Add(new Label { Text = "Double-click time (ms)", AutoSize = true }, 0, 2);
            _doubleClickSelector = new NumericUpDown
            {
                Minimum = 200,
                Maximum = 900,
                Value = SystemInformation.DoubleClickTime,
                Increment = 10,
                Dock = DockStyle.Left,
                Width = 100
            };
            pointerLayout.Controls.Add(_doubleClickSelector, 1, 2);
            pointerLayout.SetColumnSpan(_doubleClickSelector, 2);

            pointerGroup.Controls.Add(pointerLayout);
            layout.Controls.Add(pointerGroup, 0, 1);

            var scrollGroup = new GroupBox
            {
                Text = "Scroll settings",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var scrollLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8)
            };
            scrollLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            scrollLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            scrollLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            scrollLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            scrollLayout.Controls.Add(new Label { Text = "Vertical scroll lines", AutoSize = true }, 0, 0);
            _scrollLinesSelector = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = GetDefaultScrollLines(),
                Dock = DockStyle.Left,
                Width = 80
            };
            scrollLayout.Controls.Add(_scrollLinesSelector, 1, 0);

            scrollLayout.Controls.Add(new Label { Text = "Horizontal scroll characters", AutoSize = true }, 0, 1);
            _scrollCharsSelector = new NumericUpDown
            {
                Minimum = 0,
                Maximum = 100,
                Value = 3,
                Dock = DockStyle.Left,
                Width = 80
            };
            scrollLayout.Controls.Add(_scrollCharsSelector, 1, 1);

            scrollGroup.Controls.Add(scrollLayout);
            layout.Controls.Add(scrollGroup, 0, 2);

            var buttonsGroup = new GroupBox
            {
                Text = "Buttons",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var buttonsLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(8),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            _swapButtonsCheckBox = new CheckBox { Text = "Swap left/right buttons", AutoSize = true };
            buttonsLayout.Controls.Add(_swapButtonsCheckBox);
            buttonsGroup.Controls.Add(buttonsLayout);
            layout.Controls.Add(buttonsGroup, 0, 3);

            var activationGroup = new GroupBox
            {
                Text = "Activation",
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var activationLayout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                Padding = new Padding(8),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            _enableCheckBox = new CheckBox { Text = "Enable this profile", Checked = true, AutoSize = true };
            _autoApplyCheckBox = new CheckBox { Text = "Apply automatically when this device is used", Checked = true, AutoSize = true };
            _applyOnStartupCheckBox = new CheckBox { Text = "Apply on app startup", AutoSize = true };
            _enableCheckBox.CheckedChanged += (_, __) =>
            {
                _autoApplyCheckBox.Enabled = _enableCheckBox.Checked;
                _applyOnStartupCheckBox.Enabled = _enableCheckBox.Checked;
            };
            activationLayout.Controls.Add(_enableCheckBox);
            activationLayout.Controls.Add(_autoApplyCheckBox);
            activationLayout.Controls.Add(_applyOnStartupCheckBox);
            activationGroup.Controls.Add(activationLayout);
            layout.Controls.Add(activationGroup, 0, 4);

            _statusLabel = new Label { Text = "Move your device to capture its path.", AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(0, 4, 0, 0) };
            layout.Controls.Add(_statusLabel, 0, 5);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Padding = new Padding(12, 8, 12, 12),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            var okButton = new Button { Text = "Save", DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            buttons.Controls.Add(okButton);
            buttons.Controls.Add(cancelButton);

            Controls.Add(layout);
            Controls.Add(buttons);

            AcceptButton = okButton;
            CancelButton = cancelButton;

            okButton.Click += (_, __) => SaveDevice();

            if (existing != null)
            {
                _nameTextBox.Text = existing.Name;
                _pathTextBox.Text = existing.DevicePath;
                _speedSelector.Value = existing.Speed;
                _selectedColor = existing.IconColor;
                _colorPreview.BackColor = _selectedColor;
                _enableCheckBox.Checked = existing.IsEnabled;
                _autoApplyCheckBox.Checked = existing.ApplyAutomatically;
                _applyOnStartupCheckBox.Checked = existing.ApplyOnStartup;
                _enhancePrecisionCheckBox.Checked = existing.EnhancePointerPrecision;
                _scrollLinesSelector.Value = ClampToRange(existing.ScrollLines, _scrollLinesSelector.Minimum, _scrollLinesSelector.Maximum);
                _scrollCharsSelector.Value = ClampToRange(existing.ScrollChars, _scrollCharsSelector.Minimum, _scrollCharsSelector.Maximum);
                _swapButtonsCheckBox.Checked = existing.SwapButtons;
                _doubleClickSelector.Value = ClampToRange(existing.DoubleClickTime, _doubleClickSelector.Minimum, _doubleClickSelector.Maximum);
                Text = "Edit device";
            }
            else
            {
                _selectedColor = GetDefaultColor(existingColors);
                _colorPreview.BackColor = _selectedColor;
                _enhancePrecisionCheckBox.Checked = true;
            }
        }

        void BeginCapture()
        {
            _captureWindow?.Dispose();
            _captureWindow = new RawInputCaptureWindow(OnDeviceCaptured);
            _statusLabel.Text = "Listening for mouse movement... move the new device now.";
        }

        void OnDeviceCaptured(string devicePath)
        {
            if (string.IsNullOrWhiteSpace(devicePath))
                return;

            _pathTextBox.Text = devicePath;
            _statusLabel.Text = "Captured device path. You can rename it and choose a default speed.";
            _captureWindow?.Dispose();
            _captureWindow = null;
            Program.EnsureRawInputRegistration();
        }

        void TestSpeed()
        {
            Program.NativeMethods.SetMouseSpeed((int)_speedSelector.Value);
            _statusLabel.Text = $"Applied speed {(int)_speedSelector.Value} for testing.";
        }

        void ChooseColor()
        {
            using (var dialog = new ColorDialog { Color = _selectedColor })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedColor = dialog.Color;
                    _colorPreview.BackColor = _selectedColor;
                }
            }
        }

        void SaveDevice()
        {
            if (string.IsNullOrWhiteSpace(_pathTextBox.Text))
            {
                MessageBox.Show(this, "Please capture or enter a device path before saving.", "Missing device path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            if (string.IsNullOrWhiteSpace(_nameTextBox.Text))
            {
                MessageBox.Show(this, "Please provide a name for the device.", "Missing name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                DialogResult = DialogResult.None;
                return;
            }

            bool isEnabled = _enableCheckBox.Checked;
            bool autoApply = isEnabled && _autoApplyCheckBox.Checked;
            bool applyOnStartup = isEnabled && _applyOnStartupCheckBox.Checked;

            NewProfile = new Program.DeviceProfile(
                _nameTextBox.Text.Trim(),
                _pathTextBox.Text.Trim(),
                (int)_speedSelector.Value,
                _selectedColor,
                isEnabled,
                autoApply,
                applyOnStartup,
                _enhancePrecisionCheckBox.Checked,
                (int)_scrollLinesSelector.Value,
                (int)_scrollCharsSelector.Value,
                _swapButtonsCheckBox.Checked,
                (int)_doubleClickSelector.Value);
        }

        Color GetDefaultColor(IEnumerable<Color> existingColors)
        {
            var palette = new List<Color>
            {
                Color.SteelBlue,
                Color.ForestGreen,
                Color.OrangeRed,
                Color.SlateBlue,
                Color.DarkCyan,
                Color.Goldenrod,
                Color.Crimson,
                Color.Teal
            };

            var existingSet = new HashSet<Color>(existingColors ?? Enumerable.Empty<Color>());
            var unused = palette.FirstOrDefault(c => !existingSet.Contains(c));
            return unused == default ? palette[0] : unused;
        }

        decimal ClampToRange(int value, decimal min, decimal max)
        {
            if (value < min)
                return min;
            if (value > max)
                return max;
            return value;
        }

        decimal GetDefaultScrollLines()
        {
            int lines = SystemInformation.MouseWheelScrollLines;
            return lines > 0 ? lines : 3;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _captureWindow?.Dispose();
            Program.EnsureRawInputRegistration();
            base.OnFormClosed(e);
        }

        class RawInputCaptureWindow : NativeWindow, IDisposable
        {
            readonly Action<string> _callback;

            public RawInputCaptureWindow(Action<string> callback)
            {
                _callback = callback;
                CreateHandle(new CreateParams());

                Program.RAWINPUTDEVICE[] rid = new[]
                {
                    new Program.RAWINPUTDEVICE
                    {
                        usUsagePage = 0x01,
                        usUsage = 0x02,
                        dwFlags = Program.NativeMethods.RIDEV_INPUTSINK,
                        hwndTarget = Handle
                    }
                };

                Program.NativeMethods.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTDEVICE)));
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == Program.NativeMethods.WM_INPUT)
                {
                    string devicePath = GetDeviceNameFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _callback(devicePath);
                }

                base.WndProc(ref m);
            }

            string GetDeviceNameFromMessage(IntPtr lParam)
            {
                uint size = 0;
                Program.NativeMethods.GetRawInputData(lParam, Program.NativeMethods.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTHEADER)));
                if (size == 0)
                    return string.Empty;

                IntPtr buffer = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (Program.NativeMethods.GetRawInputData(lParam, Program.NativeMethods.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf(typeof(Program.RAWINPUTHEADER))) != size)
                        return string.Empty;

                    Program.RAWINPUT raw = (Program.RAWINPUT)Marshal.PtrToStructure(buffer, typeof(Program.RAWINPUT));
                    if (raw.header.dwType != Program.NativeMethods.RIM_TYPEMOUSE)
                        return string.Empty;

                    if (raw.data.lLastX == 0 && raw.data.lLastY == 0)
                        return string.Empty;

                    return GetDeviceName(raw.header.hDevice);
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }
            }

            string GetDeviceName(IntPtr hDevice)
            {
                uint size = 0;
                Program.NativeMethods.GetRawInputDeviceInfo(hDevice, Program.NativeMethods.RIDI_DEVICENAME, IntPtr.Zero, ref size);
                if (size == 0)
                    return string.Empty;

                IntPtr data = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (Program.NativeMethods.GetRawInputDeviceInfo(hDevice, Program.NativeMethods.RIDI_DEVICENAME, data, ref size) == uint.MaxValue)
                        return string.Empty;

                    return Marshal.PtrToStringAnsi(data);
                }
                finally
                {
                    Marshal.FreeHGlobal(data);
                }
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }
    }
}
