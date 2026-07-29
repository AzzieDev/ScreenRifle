# ScreenRifle

A lightweight background utility that automatically captures and saves screenshots at 5 second intervals.

## Description

This 32-bit Windows Application was written in Visual Basic 6 on September 12th, 2017, as a proof-of-concept for building an unobtrusive background tool that runs solely in the Windows System Tray.

I wrote this application to run behind Respondus Lockdown Browser to gather evidence if things go wrong during my college exams if necessary, and it absolutely delivered when they did.

## Compatibility and High DPI Environments

**ScreenRifle (twinBASIC)**
No configuration is required. The twinBASIC executable includes a modern manifest and is natively `PER_MONITOR_DPI_AWARE`. It will automatically capture full-resolution screenshots on 4K and high-DPI displays right out of the box.

**ScreenRifle (VB6)**
The original application runs perfectly on modern Windows 11 systems, but because it is a legacy VB6 application, it is not natively DPI-aware. To capture full-resolution screenshots on high-DPI displays with screen scaling using this version, you must adjust the executable's properties:

1. Right-click `ScreenRifle (VB6).exe` and select **Properties**.
2. Go to the **Compatibility** tab and click **Change high DPI settings**.
3. Check the box for **Override high DPI scaling behavior**.
4. Ensure the dropdown is set to **Application**.
