using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace MultiMouseSensitivityChanger
{
    class ManageDevicesForm : Form
    {
        readonly BindingList<Program.DeviceProfile> _profiles;
        readonly ListView _listView;

        public IReadOnlyList<Program.DeviceProfile> UpdatedProfiles => _profiles.ToList();

        public ManageDevicesForm(IEnumerable<Program.DeviceProfile> profiles)
        {
            _profiles = new BindingList<Program.DeviceProfile>(profiles.Select(CloneProfile).ToList());

            Text = "Manage devices";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            ClientSize = new Size(620, 320);

            _listView = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                MultiSelect = false,
                HideSelection = false
            };
            _listView.Columns.Add("Name", 140);
            _listView.Columns.Add("Speed", 80);
            _listView.Columns.Add("Path", 360);

            var addButton = new Button { Text = "Add", AutoSize = true };
            addButton.Click += (_, __) => AddProfile();

            var editButton = new Button { Text = "Edit", AutoSize = true };
            editButton.Click += (_, __) => EditSelectedProfile();

            var removeButton = new Button { Text = "Remove", AutoSize = true };
            removeButton.Click += (_, __) => RemoveSelectedProfile();

            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Bottom,
                Padding = new Padding(10)
            };
            buttons.Controls.Add(addButton);
            buttons.Controls.Add(editButton);
            buttons.Controls.Add(removeButton);

            var confirmButtons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Bottom,
                Padding = new Padding(10)
            };
            var okButton = new Button { Text = "Save", DialogResult = DialogResult.OK, AutoSize = true };
            var cancelButton = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true };
            confirmButtons.Controls.Add(okButton);
            confirmButtons.Controls.Add(cancelButton);
            AcceptButton = okButton;
            CancelButton = cancelButton;

            Controls.Add(_listView);
            Controls.Add(buttons);
            Controls.Add(confirmButtons);

            RefreshList();
        }

        static Program.DeviceProfile CloneProfile(Program.DeviceProfile source)
        {
            return new Program.DeviceProfile(source.Name, source.DevicePath, source.Speed, source.IconColor);
        }

        void RefreshList()
        {
            _listView.Items.Clear();
            foreach (var profile in _profiles)
            {
                var item = new ListViewItem(profile.Name)
                {
                    Tag = profile
                };
                item.SubItems.Add(profile.Speed.ToString());
                item.SubItems.Add(profile.DevicePath);
                item.BackColor = profile.IconColor;
                item.ForeColor = profile.IconColor.GetBrightness() < 0.5 ? Color.White : Color.Black;
                _listView.Items.Add(item);
            }
        }

        void AddProfile()
        {
            using (var dialog = new AddDeviceForm(null, NameExists))
            {
                if (dialog.ShowDialog() == DialogResult.OK && dialog.NewProfile != null)
                {
                    _profiles.Add(dialog.NewProfile);
                    RefreshList();
                }
            }
        }

        void EditSelectedProfile()
        {
            var selected = _listView.SelectedItems.Count > 0 ? _listView.SelectedItems[0].Tag as Program.DeviceProfile : null;
            if (selected == null)
                return;

            using (var dialog = new AddDeviceForm(selected, name => NameExists(name, selected)))
            {
                if (dialog.ShowDialog() == DialogResult.OK && dialog.NewProfile != null)
                {
                    selected.Name = dialog.NewProfile.Name;
                    selected.DevicePath = dialog.NewProfile.DevicePath;
                    selected.Speed = dialog.NewProfile.Speed;
                    selected.IconColor = dialog.NewProfile.IconColor;
                    RefreshList();
                }
            }
        }

        void RemoveSelectedProfile()
        {
            var selected = _listView.SelectedItems.Count > 0 ? _listView.SelectedItems[0].Tag as Program.DeviceProfile : null;
            if (selected == null)
                return;

            var confirm = MessageBox.Show(this, $"Remove device '{selected.Name}'?", "Confirm removal", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm == DialogResult.Yes)
            {
                _profiles.Remove(selected);
                RefreshList();
            }
        }

        bool NameExists(string name, Program.DeviceProfile ignore = null)
        {
            return _profiles.Any(p => !ReferenceEquals(p, ignore) && p.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
    }
}
