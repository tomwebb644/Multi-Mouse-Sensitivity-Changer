using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class ManageDevicesForm : Form
    {
        readonly ListView _deviceList;
        readonly Button _editButton;
        readonly Button _removeButton;
        public List<Program.DeviceProfile> Devices { get; }

        public ManageDevicesForm(IEnumerable<Program.DeviceProfile> devices)
        {
            Text = "Manage devices";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(980, 420);

            Devices = devices?.Select(CloneProfile).ToList() ?? new List<Program.DeviceProfile>();

            _deviceList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false
            };
            _deviceList.Columns.Add("Name", 140);
            _deviceList.Columns.Add("Speed", 60);
            _deviceList.Columns.Add("Precision", 80);
            _deviceList.Columns.Add("V Scroll", 70);
            _deviceList.Columns.Add("H Scroll", 70);
            _deviceList.Columns.Add("Swap", 60);
            _deviceList.Columns.Add("Startup", 70);
            _deviceList.Columns.Add("Auto", 60);
            _deviceList.Columns.Add("Icon color", 120);
            _deviceList.Columns.Add("Device path", 300);
            _deviceList.SelectedIndexChanged += (_, __) => UpdateButtons();

            var addButton = new Button { Text = "Add", AutoSize = true };
            addButton.Click += (_, __) => AddDevice();

            _editButton = new Button { Text = "Edit", AutoSize = true };
            _editButton.Click += (_, __) => EditSelected();

            _removeButton = new Button { Text = "Remove", AutoSize = true };
            _removeButton.Click += (_, __) => RemoveSelected();

            var closeButton = new Button { Text = "Close", AutoSize = true, DialogResult = DialogResult.OK };

            var buttons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0),
                AutoSize = true
            };
            buttons.Controls.Add(closeButton);
            buttons.Controls.Add(_removeButton);
            buttons.Controls.Add(_editButton);
            buttons.Controls.Add(addButton);

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(12)
            };
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(_deviceList, 0, 0);
            var buttonsContainer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 8, 0, 0)
            };
            buttonsContainer.Controls.Add(buttons);
            layout.Controls.Add(buttonsContainer, 0, 1);

            Controls.Add(layout);

            AcceptButton = closeButton;
            CancelButton = closeButton;

            RefreshList();
        }

        void RefreshList()
        {
            _deviceList.Items.Clear();
            foreach (var device in Devices)
            {
                var item = new ListViewItem(device.Name)
                {
                    Tag = device
                };
                item.SubItems.Add(device.Speed.ToString());
                item.SubItems.Add(device.EnhancePointerPrecision ? "On" : "Off");
                item.SubItems.Add(device.VerticalScrollLines.ToString());
                item.SubItems.Add(device.HorizontalScrollChars.ToString());
                item.SubItems.Add(device.SwapButtons ? "Yes" : "No");
                item.SubItems.Add(device.ApplyOnStartup ? "Yes" : "No");
                item.SubItems.Add(device.AutoApply ? "Yes" : "No");
                item.SubItems.Add(device.IconColor.Name);
                item.SubItems.Add(device.DevicePath);
                _deviceList.Items.Add(item);
            }

            UpdateButtons();
        }

        void UpdateButtons()
        {
            bool hasSelection = _deviceList.SelectedItems.Count > 0;
            _editButton.Enabled = hasSelection;
            _removeButton.Enabled = hasSelection;
        }

        void AddDevice()
        {
            using (var form = new AddDeviceForm(existingColors: Devices.Select(d => d.IconColor)))
            {
                if (form.ShowDialog() == DialogResult.OK && form.NewProfile != null)
                {
                    Devices.Add(CloneProfile(form.NewProfile));
                    RefreshList();
                }
            }
        }

        void EditSelected()
        {
            if (_deviceList.SelectedItems.Count == 0)
                return;

            var profile = _deviceList.SelectedItems[0].Tag as Program.DeviceProfile;
            if (profile == null)
                return;

            using (var form = new AddDeviceForm(profile, Devices.Where(d => !ReferenceEquals(d, profile)).Select(d => d.IconColor)))
            {
                if (form.ShowDialog() == DialogResult.OK && form.NewProfile != null)
                {
                    profile.Name = form.NewProfile.Name;
                    profile.DevicePath = form.NewProfile.DevicePath;
                    profile.Speed = form.NewProfile.Speed;
                    profile.IconColor = form.NewProfile.IconColor;
                    profile.EnhancePointerPrecision = form.NewProfile.EnhancePointerPrecision;
                    profile.VerticalScrollLines = form.NewProfile.VerticalScrollLines;
                    profile.HorizontalScrollChars = form.NewProfile.HorizontalScrollChars;
                    profile.SwapButtons = form.NewProfile.SwapButtons;
                    profile.ApplyOnStartup = form.NewProfile.ApplyOnStartup;
                    profile.AutoApply = form.NewProfile.AutoApply;
                    RefreshList();
                }
            }
        }

        void RemoveSelected()
        {
            if (_deviceList.SelectedItems.Count == 0)
                return;

            var profile = _deviceList.SelectedItems[0].Tag as Program.DeviceProfile;
            if (profile == null)
                return;

            Devices.Remove(profile);
            RefreshList();
        }

        Program.DeviceProfile CloneProfile(Program.DeviceProfile profile)
        {
            return new Program.DeviceProfile(
                profile.Name,
                profile.DevicePath,
                profile.Speed,
                profile.IconColor,
                profile.EnhancePointerPrecision,
                profile.VerticalScrollLines,
                profile.HorizontalScrollChars,
                profile.SwapButtons,
                profile.ApplyOnStartup,
                profile.AutoApply);
        }
    }
}
