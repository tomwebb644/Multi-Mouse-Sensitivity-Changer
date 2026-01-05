using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class AddDeviceForm : Form
    {
        readonly TextBox _nameTextBox;
        readonly TextBox _pathTextBox;
        readonly NumericUpDown _speedSelector;
        readonly Label _statusLabel;
        readonly Panel _colorPreview;
        Color _selectedColor = Color.Gray;
        RawInputCaptureWindow _captureWindow;

        public Program.DeviceProfile NewProfile { get; private set; }

        public AddDeviceForm(Program.DeviceProfile existing = null)
        {
            Text = "Add new device";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(420, 260);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 7,
                Padding = new Padding(10),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 70));

            layout.Controls.Add(new Label { Text = "Step 1: Capture the device path", AutoSize = true }, 0, 0);
            _pathTextBox = new TextBox { Dock = DockStyle.Fill };
            layout.Controls.Add(_pathTextBox, 1, 0);
            var captureButton = new Button { Text = "Capture from movement", Dock = DockStyle.Left };
            captureButton.Click += (_, __) => BeginCapture();
            layout.Controls.Add(captureButton, 1, 1);

            layout.Controls.Add(new Label { Text = "Step 2: Name", AutoSize = true }, 0, 2);
            _nameTextBox = new TextBox { Dock = DockStyle.Fill };
            layout.Controls.Add(_nameTextBox, 1, 2);

            layout.Controls.Add(new Label { Text = "Step 3: Default speed", AutoSize = true }, 0, 3);
            _speedSelector = new NumericUpDown
            {
                Minimum = 1,
                Maximum = 20,
                Value = 10,
                Dock = DockStyle.Left,
                Width = 80
            };
            layout.Controls.Add(_speedSelector, 1, 3);
            var testButton = new Button { Text = "Test speed", Dock = DockStyle.Left };
            testButton.Click += (_, __) => TestSpeed();
            layout.Controls.Add(testButton, 1, 4);

            layout.Controls.Add(new Label { Text = "Icon color", AutoSize = true }, 0, 5);
            var colorButton = new Button { Text = "Choose color", Dock = DockStyle.Left };
            colorButton.Click += (_, __) => ChooseColor();
            var colorPanel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, Dock = DockStyle.Fill };
            _colorPreview = new Panel { Width = 32, Height = 32, BackColor = _selectedColor, BorderStyle = BorderStyle.FixedSingle };
            colorPanel.Controls.Add(colorButton);
            colorPanel.Controls.Add(_colorPreview);
            layout.Controls.Add(colorPanel, 1, 5);

            _statusLabel = new Label { Text = "Move your device to capture its path.", AutoSize = true, ForeColor = Color.DimGray };
            layout.Controls.Add(_statusLabel, 0, 6);
            layout.SetColumnSpan(_statusLabel, 2);

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Padding = new Padding(10)
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
                Text = "Edit device";
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

            NewProfile = new Program.DeviceProfile(_nameTextBox.Text.Trim(), _pathTextBox.Text.Trim(), (int)_speedSelector.Value, _selectedColor);
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _captureWindow?.Dispose();
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
