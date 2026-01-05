// Program.cs
// .NET Framework 4.7.2
// Auto-switch Windows pointer speed based on which Raw Input mouse device last moved.

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

class Program
{
    // ====== SET THESE AFTER FIRST RUN ======
    static string X8_DEVICE_PATH =
    @"\\?\HID#VID_1997&PID_2433&MI_01&Col01#9&354159c2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

    static string MOUSE_DEVICE_PATH =
    @"\\?\HID#VID_046D&PID_C539&MI_01&Col01#8&10c9e4b2&0&0000#{378de44c-56ef-11d1-bc8c-00a0c91405dd}";

    static int X8_SPEED = 20;    // 1..20
    static int MOUSE_SPEED = 10; // 1..20

    static int MIN_SWITCH_MS = 200;
    // ======================================

    static IntPtr _hwnd;
    static int _lastSpeed = -1;
    static long _lastSwitchMs = 0;
    static Stopwatch _sw = Stopwatch.StartNew();

    static void Main()
    {
        WNDCLASS wc = new WNDCLASS();
        wc.lpszClassName = "RawInputSpeedSwitcherWindow";
        wc.lpfnWndProc = WndProc;

        ushort atom = RegisterClass(ref wc);
        if (atom == 0) ThrowLastWin32("RegisterClass failed");

        _hwnd = CreateWindowEx(
            0, wc.lpszClassName, "RawInputSpeedSwitcher",
            0, 0, 0, 0, 0,
            HWND_MESSAGE, IntPtr.Zero, IntPtr.Zero, IntPtr.Zero);

        if (_hwnd == IntPtr.Zero) ThrowLastWin32("CreateWindowEx failed");

        RAWINPUTDEVICE[] rid = new RAWINPUTDEVICE[]
        {
            new RAWINPUTDEVICE
            {
                usUsagePage = 0x01,
                usUsage = 0x02,
                dwFlags = RIDEV_INPUTSINK,
                hwndTarget = _hwnd
            }
        };

        if (!RegisterRawInputDevices(rid, (uint)rid.Length, (uint)Marshal.SizeOf(typeof(RAWINPUTDEVICE))))
            ThrowLastWin32("RegisterRawInputDevices failed");

        Console.WriteLine("Raw Input Speed Switcher running.");
        Console.WriteLine("Move devices to see their Raw Input device paths.");
        Console.WriteLine();

        MSG msg;
        while (GetMessage(out msg, IntPtr.Zero, 0, 0) > 0)
        {
            TranslateMessage(ref msg);
            DispatchMessage(ref msg);
        }
    }

    static IntPtr WndProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam)
    {
        if (msg == WM_INPUT)
            HandleRawInput(lParam);

        return DefWindowProc(hWnd, msg, wParam, lParam);
    }

    static void HandleRawInput(IntPtr hRawInput)
    {
        uint size = 0;
        GetRawInputData(hRawInput, RID_INPUT, IntPtr.Zero, ref size,
            (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER)));

        if (size == 0) return;

        IntPtr buffer = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputData(hRawInput, RID_INPUT, buffer, ref size,
                (uint)Marshal.SizeOf(typeof(RAWINPUTHEADER))) != size)
                return;

            RAWINPUT raw = (RAWINPUT)Marshal.PtrToStructure(buffer, typeof(RAWINPUT));
            if (raw.header.dwType != RIM_TYPEMOUSE) return;

            if (raw.data.lLastX == 0 && raw.data.lLastY == 0) return;

            string devicePath = GetDeviceName(raw.header.hDevice);
            if (string.IsNullOrEmpty(devicePath)) return;

            Console.WriteLine("Device: " + devicePath);

            if (string.IsNullOrEmpty(X8_DEVICE_PATH) || string.IsNullOrEmpty(MOUSE_DEVICE_PATH))
                return;

            int targetSpeed =
                devicePath.Equals(X8_DEVICE_PATH, StringComparison.OrdinalIgnoreCase) ? X8_SPEED :
                devicePath.Equals(MOUSE_DEVICE_PATH, StringComparison.OrdinalIgnoreCase) ? MOUSE_SPEED :
                -1;

            if (targetSpeed < 0) return;

            long now = _sw.ElapsedMilliseconds;
            if (targetSpeed != _lastSpeed && (now - _lastSwitchMs) >= MIN_SWITCH_MS)
            {
                SetMouseSpeed(targetSpeed);
                _lastSpeed = targetSpeed;
                _lastSwitchMs = now;
                Console.WriteLine("-> Speed " + targetSpeed);
            }
        }
        finally
        {
            Marshal.FreeHGlobal(buffer);
        }
    }

    static string GetDeviceName(IntPtr hDevice)
    {
        uint size = 0;
        GetRawInputDeviceInfo(hDevice, RIDI_DEVICENAME, IntPtr.Zero, ref size);
        if (size == 0) return "";

        IntPtr data = Marshal.AllocHGlobal((int)size);
        try
        {
            if (GetRawInputDeviceInfo(hDevice, RIDI_DEVICENAME, data, ref size) == uint.MaxValue)
                return "";

            return Marshal.PtrToStringAnsi(data);
        }
        finally
        {
            Marshal.FreeHGlobal(data);
        }
    }

    static void SetMouseSpeed(int speed)
    {
        if (speed < 1) speed = 1;
        if (speed > 20) speed = 20;
        SystemParametersInfo(SPI_SETMOUSESPEED, 0, (IntPtr)speed, SPIF_SENDCHANGE);
    }

    static void ThrowLastWin32(string msg)
    {
        throw new Exception(msg + " Win32=" + Marshal.GetLastWin32Error());
    }

    // ================= Win32 =================

    const uint WM_INPUT = 0x00FF;
    const uint RID_INPUT = 0x10000003;
    const uint RIM_TYPEMOUSE = 0;
    const uint RIDI_DEVICENAME = 0x20000007;
    const uint RIDEV_INPUTSINK = 0x00000100;
    const uint SPI_SETMOUSESPEED = 0x0071;
    const uint SPIF_SENDCHANGE = 0x02;

    static readonly IntPtr HWND_MESSAGE = new IntPtr(-3);

    [StructLayout(LayoutKind.Sequential)]
    struct WNDCLASS
    {
        public uint style;
        public WndProcDelegate lpfnWndProc;
        public int cbClsExtra;
        public int cbWndExtra;
        public IntPtr hInstance;
        public IntPtr hIcon;
        public IntPtr hCursor;
        public IntPtr hbrBackground;
        public string lpszMenuName;
        public string lpszClassName;
    }

    delegate IntPtr WndProcDelegate(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [StructLayout(LayoutKind.Sequential)]
    struct MSG
    {
        public IntPtr hwnd;
        public uint message;
        public IntPtr wParam;
        public IntPtr lParam;
        public uint time;
        public POINT pt;
        public uint lPrivate;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINT { public int x, y; }

    [StructLayout(LayoutKind.Sequential)]
    struct RAWINPUTDEVICE
    {
        public ushort usUsagePage;
        public ushort usUsage;
        public uint dwFlags;
        public IntPtr hwndTarget;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct RAWINPUTHEADER
    {
        public uint dwType;
        public uint dwSize;
        public IntPtr hDevice;
        public IntPtr wParam;
    }

    [StructLayout(LayoutKind.Explicit)]
    struct RAWINPUT
    {
        [FieldOffset(0)] public RAWINPUTHEADER header;
        [FieldOffset(16)] public RAWMOUSE data;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct RAWMOUSE
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

    [DllImport("user32.dll", SetLastError = true)]
    static extern ushort RegisterClass(ref WNDCLASS lpWndClass);

    [DllImport("user32.dll", SetLastError = true)]
    static extern IntPtr CreateWindowEx(
        uint dwExStyle, string lpClassName, string lpWindowName, uint dwStyle,
        int x, int y, int nWidth, int nHeight,
        IntPtr hWndParent, IntPtr hMenu, IntPtr hInstance, IntPtr lpParam);

    [DllImport("user32.dll")]
    static extern IntPtr DefWindowProc(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool RegisterRawInputDevices(
        RAWINPUTDEVICE[] pRawInputDevices, uint uiNumDevices, uint cbSize);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetRawInputData(
        IntPtr hRawInput, uint uiCommand, IntPtr pData, ref uint pcbSize, uint cbSizeHeader);

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint GetRawInputDeviceInfo(
        IntPtr hDevice, uint uiCommand, IntPtr pData, ref uint pcbSize);

    [DllImport("user32.dll")]
    static extern sbyte GetMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax);

    [DllImport("user32.dll")]
    static extern bool TranslateMessage(ref MSG lpMsg);

    [DllImport("user32.dll")]
    static extern IntPtr DispatchMessage(ref MSG lpmsg);

    [DllImport("user32.dll", SetLastError = true)]
    static extern bool SystemParametersInfo(
        uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
}