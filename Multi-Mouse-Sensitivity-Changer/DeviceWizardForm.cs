using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class DeviceWizardForm : Form
    {
        readonly HashSet<string> _existingPaths;
        readonly Label _instructionLabel;
        readonly TextBox _nameBox;
        readonly TextBox _pathBox;
        readonly NumericUpDown _speedUpDown;
        readonly Label _statusLabel;
        readonly Button _captureButton;
        readonly Button _testButton;
        readonly Button _saveButton;
        RawInputCaptureWindow _captureWindow;
        bool _capturing;

        public string DeviceName => _nameBox.Text.Trim();
        public string DevicePath => _pathBox.Text.Trim();
        public int DeviceSpeed => (int)_speedUpDown.Value;

        public DeviceWizardForm(IEnumerable<string> existingPaths)
        {
            _existingPaths = new HashSet<string>(existingPaths ?? Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);

            Text = "Add new device";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(420, 260);

            _instructionLabel = new Label
            {
                Text = "Follow the steps to add a device:\n  1) Click 'Start capture' and move the device.\n  2) Give it a friendly name.\n  3) Pick a default speed and test it.",
                AutoSize = false,
                Size = new Size(390, 70),
                Location = new Point(12, 10)
            };

            var nameLabel = new Label { Text = "Device name", Location = new Point(12, 90), AutoSize = true };
            _nameBox = new TextBox { Location = new Point(110, 86), Width = 280 };

            var pathLabel = new Label { Text = "Device path", Location = new Point(12, 120), AutoSize = true };
            _pathBox = new TextBox { Location = new Point(110, 116), Width = 280 };

            _captureButton = new Button { Text = "Start capture", Location = new Point(110, 146), Width = 110 };
            _captureButton.Click += (_, __) => ToggleCapture();

            var speedLabel = new Label { Text = "Default speed", Location = new Point(12, 180), AutoSize = true };
            _speedUpDown = new NumericUpDown { Location = new Point(110, 176), Minimum = 1, Maximum = 20, Value = 10 };

            _testButton = new Button { Text = "Test speed", Location = new Point(200, 174), Width = 90 };
            _testButton.Click += (_, __) => TestSpeed();

            _statusLabel = new Label { Location = new Point(12, 210), AutoSize = true, ForeColor = Color.Gray };

            _saveButton = new Button { Text = "Save", Location = new Point(240, 220), Width = 70 };
            _saveButton.Click += (_, __) => OnSave();

            var cancelButton = new Button { Text = "Cancel", Location = new Point(320, 220), Width = 70, DialogResult = DialogResult.Cancel };

            Controls.AddRange(new Control[]
            {
                _instructionLabel, nameLabel, _nameBox, pathLabel, _pathBox, _captureButton,
                speedLabel, _speedUpDown, _testButton, _statusLabel, _saveButton, cancelButton
            });
        }

        void ToggleCapture()
        {
            if (_capturing)
            {
                StopCapture("Capture stopped.");
                return;
            }

            _statusLabel.Text = "Move the device you want to add...";
            _statusLabel.ForeColor = Color.DarkGreen;
            _captureButton.Text = "Stop capture";
            _capturing = true;
            _captureWindow = new RawInputCaptureWindow(OnDeviceCaptured);
        }

        void StopCapture(string status)
        {
            _captureWindow?.Dispose();
            _captureWindow = null;
            _capturing = false;
            _captureButton.Text = "Start capture";
            _statusLabel.Text = status;
            _statusLabel.ForeColor = Color.Gray;
        }

        void OnDeviceCaptured(string devicePath)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new Action<string>(OnDeviceCaptured), devicePath);
                return;
            }

            StopCapture("Device captured.");
            _pathBox.Text = devicePath;
        }

        void TestSpeed()
        {
            NativeMethods.SetMouseSpeed(DeviceSpeed);
            _statusLabel.Text = $"Applied test speed {DeviceSpeed}.";
            _statusLabel.ForeColor = Color.Black;
        }

        void OnSave()
        {
            if (string.IsNullOrWhiteSpace(DeviceName))
            {
                MessageBox.Show(this, "Please provide a device name.", "Missing name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(DevicePath))
            {
                MessageBox.Show(this, "Please capture or enter a device path.", "Missing device", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (_existingPaths.Contains(DevicePath))
            {
                MessageBox.Show(this, "That device is already configured.", "Duplicate device", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            StopCapture(string.Empty);
            base.OnFormClosed(e);
        }

        class RawInputCaptureWindow : NativeWindow, IDisposable
        {
            readonly Action<string> _captureCallback;

            public RawInputCaptureWindow(Action<string> captureCallback)
            {
                _captureCallback = captureCallback;
                CreateHandle(new CreateParams());

                if (!Program.RawInputHelper.RegisterForMouseInput(Handle))
                    throw new InvalidOperationException("RegisterRawInputDevices failed for capture window, Win32=" + System.Runtime.InteropServices.Marshal.GetLastWin32Error());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == Program.NativeMethods.WM_INPUT)
                {
                    string devicePath = Program.RawInputHelper.GetDeviceNameFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _captureCallback(devicePath);
                }

                base.WndProc(ref m);
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }
    }
}
