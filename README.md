# Multi-Mouse Sensitivity Changer

Multi-Mouse Sensitivity Changer is a Windows tray utility that automatically switches the global mouse speed based on the device you are actively using. It is useful if you regularly swap between multiple mice with different DPI settings and want each one to feel consistent without changing drivers or profiles.

## Features
- **Per-device speed profiles.** Capture the device path from mouse movement, give it a friendly name, and assign a default Windows mouse speed (1–20). The app applies the correct speed whenever that device becomes active.
- **Tray icon controls.** The tray menu shows the active device and lets you adjust each device's speed on the fly. Speeds update immediately for the device that is in use.
- **Color-coded icons.** Each device can have its own accent color and initials, which are used to generate tray icons so you can tell which profile is active at a glance.
- **Device management dialogs.** Add devices with guided steps, test speeds, choose icon colors, or edit and remove existing profiles from the Manage Devices window.
- **Start with Windows.** Optional startup toggle so the tray app can launch automatically.
- **Persistent settings.** Profiles are stored under the current user's registry key (`Software/MultiMouseSensitivityChanger`) so they remain available between sessions.

## Building and running
1. Open `Multi-Mouse-Sensitivity-Changer.sln` in Visual Studio (Windows only).
2. Restore NuGet packages if prompted.
3. Build the solution and run the application. A tray icon will appear when the app is running.

## Usage
1. Right-click the tray icon and choose **Add new device...**.
2. Click **Capture from movement** and move the mouse you want to register. The device path will be captured automatically.
3. Provide a name, pick a default speed, and optionally choose an icon color. Use **Test speed** to try the selected value before saving.
4. Repeat for additional devices. The tray menu will show each device with its speed options. Switching mice will automatically apply the correct speed.
5. Use **Manage devices...** from the tray menu to edit or remove profiles, and **Start with Windows** to enable auto-start.

## Notes
- A short switch delay is enforced to avoid rapid toggling when devices send overlapping input.
- The app uses the Windows Raw Input API to detect which physical mouse is active.
