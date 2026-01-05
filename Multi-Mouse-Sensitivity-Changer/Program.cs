using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

namespace MultiMouseSensitivityChanger
{
    static class Program
    {
        // ====== SET THESE AFTER FIRST RUN ======
        const string SettingsKeyPath = "Software\\MultiMouseSensitivityChanger";
        static string X8_DEVICE_PATH =
        @"\\?\HID#VID_1997&PID_2433&MI_01&Col01#9&354159c2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

        static string MOUSE_DEVICE_PATH =
        @"\\?\HID#VID_046D&PID_C539&MI_01&Col01#8&10c9e4b2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

        static int X8_SPEED = 20;    // 1..20
        static int MOUSE_SPEED = 10; // 1..20

        static int MIN_SWITCH_MS = 200;
        // ======================================

        static int _lastSpeed = -1;
        static long _lastSwitchMs = 0;
        static readonly Stopwatch _sw = Stopwatch.StartNew();

        static NotifyIcon _notifyIcon;
        static ContextMenuStrip _menu;
        static ToolStripMenuItem _activeDeviceItem;
        static ToolStripMenuItem _startupItem;
        static ToolStripMenuItem _addDeviceItem;
        static readonly List<DeviceProfile> _devices = new List<DeviceProfile>();

        static Icon _mouseIcon;
        static Icon _x8Icon;
        static Icon _defaultIcon;

        static string _activeDevicePath = string.Empty;

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new TrayApplicationContext());
        }

        static void OnDeviceChanged(string devicePath)
        {
            if (string.IsNullOrEmpty(devicePath))
                return;

            DeviceProfile profile = _devices.FirstOrDefault(d => devicePath.Equals(d.DevicePath, StringComparison.OrdinalIgnoreCase));
            if (profile == null)
                return;

            int targetSpeed = profile.Speed;
            long now = _sw.ElapsedMilliseconds;
            if (targetSpeed != _lastSpeed && (now - _lastSwitchMs) >= MIN_SWITCH_MS)
            {
                NativeMethods.SetMouseSpeed(targetSpeed);
                _lastSpeed = targetSpeed;
                _lastSwitchMs = now;
                UpdateActiveDevice(profile, targetSpeed);
            }
        }

        static void LoadDevices()
        {
            int mouseSpeed = ReadSpeed("Mouse", MOUSE_SPEED);
            int x8Speed = ReadSpeed("X8", X8_SPEED);

            List<DeviceProfile> loaded = LoadDeviceProfilesFromRegistry().ToList();
            if (loaded.Count == 0)
            {
                loaded.Add(new DeviceProfile("Mouse", MOUSE_DEVICE_PATH, mouseSpeed));
                loaded.Add(new DeviceProfile("X8", X8_DEVICE_PATH, x8Speed));
            }

            _devices.Clear();
            _devices.AddRange(loaded.Select(d => new DeviceProfile(d.Name, d.DevicePath, ClampSpeed(d.Speed, MOUSE_SPEED))));
        }

        static int ReadSpeed(string deviceKey, int defaultValue)
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                object value = key?.GetValue(deviceKey);

                if (value is int intValue)
                    return ClampSpeed(intValue, defaultValue);

                if (value is string str && int.TryParse(str, out int parsed))
                    return ClampSpeed(parsed, defaultValue);
            }

            return defaultValue;
        }

        static IEnumerable<DeviceProfile> LoadDeviceProfilesFromRegistry()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsKeyPath, false))
            {
                if (key == null)
                    yield break;

                object value = key.GetValue("Devices");
                string[] entries = value as string[];
                if (entries == null)
                    yield break;

                foreach (string entry in entries)
                {
                    string[] parts = entry.Split('|');
                    if (parts.Length != 3)
                        continue;

                    if (!int.TryParse(parts[2], out int speed))
                        continue;

                    yield return new DeviceProfile(parts[0], parts[1], speed);
                }
            }
        }

        static void SaveDevices()
        {
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsKeyPath))
            {
                string[] entries = _devices
                    .Select(d => $"{d.Name}|{d.DevicePath}|{ClampSpeed(d.Speed, d.Speed)}")
                    .ToArray();

                key.SetValue("Devices", entries, RegistryValueKind.MultiString);
            }
        }

        static int ClampSpeed(int speed, int fallback)
        {
            if (speed < 1 || speed > 20)
                return fallback;

            return speed;
        }

        static void UpdateActiveDevice(DeviceProfile profile, int speed)
        {
            _activeDevicePath = profile.DevicePath;
            _activeDeviceItem.Text = $"Active: {profile.Name} (speed {speed})";

            Icon icon = profile.Name.IndexOf("X8", StringComparison.OrdinalIgnoreCase) >= 0
                ? _x8Icon
                : _mouseIcon;

            _notifyIcon.Icon = icon ?? _defaultIcon;
            _notifyIcon.Text = $"{profile.Name} speed {speed}";
        }

        static void InitializeTrayIcon()
        {
            _mouseIcon = CreateIcon(Color.DodgerBlue, "M");
            _x8Icon = CreateIcon(Color.MediumPurple, "X8");
            _defaultIcon = CreateIcon(Color.Gray, "MM");

            _menu = new ContextMenuStrip();
            BuildMenuItems();

            _notifyIcon = new NotifyIcon
            {
                Visible = true,
                Icon = _defaultIcon,
                Text = "Multi-Mouse Sensitivity Changer",
                ContextMenuStrip = _menu
            };
        }

        static void BuildMenuItems()
        {
            _menu.Items.Clear();

            _activeDeviceItem = new ToolStripMenuItem("Active: none") { Enabled = false };
            _menu.Items.Add(_activeDeviceItem);
            _menu.Items.Add(new ToolStripSeparator());

            foreach (DeviceProfile profile in _devices.OrderBy(d => d.Name))
            {
                _menu.Items.Add(BuildSpeedMenu(profile));
            }

            _menu.Items.Add(new ToolStripSeparator());

            _addDeviceItem = new ToolStripMenuItem("Add device...");
            _addDeviceItem.Click += (_, __) => ShowDeviceWizard();
            _menu.Items.Add(_addDeviceItem);

            _menu.Items.Add(new ToolStripSeparator());

            _startupItem = new ToolStripMenuItem("Start with Windows")
            {
                CheckOnClick = true,
                Checked = StartupManager.IsEnabled()
            };
            _startupItem.CheckedChanged += (_, __) => StartupManager.SetEnabled(_startupItem.Checked);
            _menu.Items.Add(_startupItem);

            _menu.Items.Add(new ToolStripSeparator());
            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => Application.Exit();
            _menu.Items.Add(exitItem);
        }

        static ToolStripMenuItem BuildSpeedMenu(DeviceProfile profile)
        {
            var menu = new ToolStripMenuItem($"{profile.Name} speed");
            for (int i = 1; i <= 20; i++)
            {
                var item = new ToolStripMenuItem(i.ToString())
                {
                    Tag = new SpeedTag(profile.DevicePath, i),
                    Checked = i == profile.Speed
                };
                item.Click += OnSpeedMenuClick;
                menu.DropDownItems.Add(item);
            }
            menu.ToolTipText = profile.DevicePath;
            return menu;
        }

        static void ShowDeviceWizard()
        {
            using (var wizard = new DeviceWizardForm(_devices.Select(d => d.DevicePath)))
            {
                if (wizard.ShowDialog() == DialogResult.OK)
                {
                    var profile = new DeviceProfile(wizard.DeviceName, wizard.DevicePath, wizard.DeviceSpeed);
                    _devices.Add(profile);
                    SaveDevices();
                    BuildMenuItems();

                    var active = _devices.FirstOrDefault(d => string.Equals(d.DevicePath, _activeDevicePath, StringComparison.OrdinalIgnoreCase));
                    if (active != null)
                        UpdateActiveDevice(active, active.Speed);
                }
            }
        }

        static void OnSpeedMenuClick(object sender, EventArgs e)
        {
            if (sender is ToolStripMenuItem item && item.Tag is SpeedTag tag)
            {
                DeviceProfile profile = _devices.FirstOrDefault(d => d.DevicePath.Equals(tag.DevicePath, StringComparison.OrdinalIgnoreCase));
                if (profile == null)
                    return;

                profile.Speed = tag.Speed;
                SaveDevices();
                ToolStripMenuItem parentMenu = item.OwnerItem as ToolStripMenuItem;
                if (parentMenu != null)
                    UpdateMenuChecks(parentMenu, tag.Speed);

                if (string.Equals(_activeDevicePath, tag.DevicePath, StringComparison.OrdinalIgnoreCase))
                {
                    NativeMethods.SetMouseSpeed(tag.Speed);
                    _lastSpeed = tag.Speed;
                    UpdateActiveDevice(profile, tag.Speed);
                }
            }
        }

        static void UpdateMenuChecks(ToolStripMenuItem menu, int activeSpeed)
        {
            foreach (ToolStripMenuItem child in menu.DropDownItems)
            {
                child.Checked = child.Text == activeSpeed.ToString();
            }
        }

        static Icon CreateIcon(Color color, string label)
        {
            using (var bmp = new Bitmap(32, 32))
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (var brush = new SolidBrush(color))
                {
                    g.FillEllipse(brush, 0, 0, 31, 31);
                }

                using (var pen = new Pen(Color.White, 2))
                {
                    g.DrawEllipse(pen, 1, 1, 29, 29);
                }

                using (var font = new Font(FontFamily.GenericSansSerif, label.Length > 2 ? 7 : 10, FontStyle.Bold))
                using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                using (var textBrush = new SolidBrush(Color.White))
                {
                    g.DrawString(label, font, textBrush, new RectangleF(0, 0, 32, 32), format);
                }

                IntPtr hIcon = bmp.GetHicon();
                try
                {
                    using (var icon = Icon.FromHandle(hIcon))
                    {
                        return (Icon)icon.Clone();
                    }
                }
                finally
                {
                    NativeMethods.DestroyIcon(hIcon);
                }
            }
        }

        class TrayApplicationContext : ApplicationContext
        {
            readonly RawInputWindow _window;

            public TrayApplicationContext()
            {
                LoadDevices();
                InitializeTrayIcon();
                _window = new RawInputWindow(OnDeviceChanged);
            }

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _notifyIcon?.Dispose();
                    _menu?.Dispose();
                    _mouseIcon?.Dispose();
                    _x8Icon?.Dispose();
                    _defaultIcon?.Dispose();
                    _window?.Dispose();
                }
                base.Dispose(disposing);
            }
        }

        class RawInputWindow : NativeWindow, IDisposable
        {
            readonly Action<string> _deviceCallback;

            public RawInputWindow(Action<string> deviceCallback)
            {
                _deviceCallback = deviceCallback;
                CreateHandle(new CreateParams());

                if (!RawInputHelper.RegisterForMouseInput(Handle))
                    throw new InvalidOperationException("RegisterRawInputDevices failed, Win32=" + Marshal.GetLastWin32Error());
            }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == NativeMethods.WM_INPUT)
                {
                    string devicePath = RawInputHelper.GetDeviceNameFromMessage(m.LParam);
                    if (!string.IsNullOrEmpty(devicePath))
                        _deviceCallback(devicePath);
                }

                base.WndProc(ref m);
            }

            public void Dispose()
            {
                DestroyHandle();
            }
        }

        public static class RawInputHelper
        {
            public static bool RegisterForMouseInput(IntPtr handle)
            {
                RAWINPUTDEVICE[] rid =
                {
                    new RAWINPUTDEVICE
                    {
                        usUsagePage = 0x01,
                        usUsage = 0x02,
                        dwFlags = NativeMethods.RIDEV_INPUTSINK,
                        hwndTarget = handle
                    }
                };

                return NativeMethods.RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE)));
            }

            public static string GetDeviceNameFromMessage(IntPtr lParam)
            {
                uint size = 0;
                NativeMethods.GetRawInputData(lParam, NativeMethods.RID_INPUT, IntPtr.Zero, ref size, (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER)));
                if (size == 0)
                    return string.Empty;

                IntPtr buffer = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (NativeMethods.GetRawInputData(lParam, NativeMethods.RID_INPUT, buffer, ref size, (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER))) != size)
                        return string.Empty;

                    RAWINPUT raw = (RAWINPUT)Marshal.PtrToStructure(buffer, typeof(RAWINPUT));
                    if (raw.header.dwType != NativeMethods.RIM_TYPEMOUSE)
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

            static string GetDeviceName(IntPtr hDevice)
            {
                uint size = 0;
                NativeMethods.GetRawInputDeviceInfo(hDevice, NativeMethods.RIDI_DEVICENAME, IntPtr.Zero, ref size);
                if (size == 0)
                    return string.Empty;

                IntPtr data = Marshal.AllocHGlobal((int)size);
                try
                {
                    if (NativeMethods.GetRawInputDeviceInfo(hDevice, NativeMethods.RIDI_DEVICENAME, data, ref size) == uint.MaxValue)
                        return string.Empty;

                    return Marshal.PtrToStringAnsi(data);
                }
                finally
                {
                    Marshal.FreeHGlobal(data);
                }
            }
        }

        class SpeedTag
        {
            public string DevicePath { get; }
            public int Speed { get; }

            public SpeedTag(string devicePath, int speed)
            {
                DevicePath = devicePath;
                Speed = speed;
            }
        }

        class DeviceProfile
        {
            public string Name { get; set; }
            public string DevicePath { get; set; }
            public int Speed { get; set; }

            public DeviceProfile(string name, string devicePath, int speed)
            {
                Name = name;
                DevicePath = devicePath;
                Speed = speed;
            }
        }

        static class StartupManager
        {
            const string RunKeyPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Run";
            static readonly string AppName = "Multi-Mouse-Sensitivity-Changer";

            public static bool IsEnabled()
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, false))
                {
                    string value = key?.GetValue(AppName) as string;
                    return !string.IsNullOrEmpty(value);
                }
            }

            public static void SetEnabled(bool enabled)
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(RunKeyPath, true) ?? Registry.CurrentUser.CreateSubKey(RunKeyPath))
                {
                    if (enabled)
                    {
                        string exePath = Process.GetCurrentProcess().MainModule.FileName;
                        key.SetValue(AppName, '"' + exePath + '"');
                    }
                    else
                    {
                        key.DeleteValue(AppName, false);
                    }
                }
            }
        }

        internal static class NativeMethods
        {
            public const uint WM_INPUT = 0x00FF;
            public const uint RID_INPUT = 0x10000003;
            public const uint RIM_TYPEMOUSE = 0;
            public const uint RIDI_DEVICENAME = 0x20000007;
            public const uint RIDEV_INPUTSINK = 0x00000100;
            public const uint SPI_SETMOUSESPEED = 0x0071;
            public const uint SPIF_SENDCHANGE = 0x02;

            public static void SetMouseSpeed(int speed)
            {
                if (speed < 1) speed = 1;
                if (speed > 20) speed = 20;
                SystemParametersInfo(SPI_SETMOUSESPEED, 0, (IntPtr)speed, SPIF_SENDCHANGE);
            }

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern bool RegisterRawInputDevices(RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint GetRawInputData(IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

            [DllImport("user32.dll", SetLastError = true)]
            public static extern uint GetRawInputDeviceInfo(IntPtr hDevice, uint uiCommand, IntPtr pData, ref uint pcbSize);

            [DllImport("user32.dll")]
            public static extern bool DestroyIcon(IntPtr handle);
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWINPUTDEVICE
        {
            public ushort usUsagePage;
            public ushort usUsage;
            public uint dwFlags;
            public IntPtr hwndTarget;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWINPUTHEADER
        {
            public uint dwType;
            public uint dwSize;
            public IntPtr hDevice;
            public IntPtr wParam;
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct RAWINPUT
        {
            [FieldOffset(0)] public RAWINPUTHEADER header;
            [FieldOffset(16)] public RAWMOUSE data;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RAWMOUSE
        {
            public ushort usFlags;
            public uint ulButtons;
            public ushort usButtonFlags;
            public ushort usButtonData;
            public uint ulRawButtons;
            public int lLastX;
            public int lLastY;
            public uint ulExtraInformation;
        }
    }
}
