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
            ClientSize = new Size(720, 320);

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
            _deviceList.Columns.Add("Icon color", 120);
            _deviceList.Columns.Add("Device path", 360);
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
            return new Program.DeviceProfile(profile.Name, profile.DevicePath, profile.Speed, profile.IconColor);
        }
    }
}
