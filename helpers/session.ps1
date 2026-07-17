function Disable-ConsoleQuickEdit {
    if (-not $IsWindows) {
        return
    }

    if (-not ('Native.ConsoleMode' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Native
{
    public static class ConsoleMode
    {
        private const int STD_INPUT_HANDLE = -10;

        private const uint ENABLE_QUICK_EDIT_MODE = 0x0040;
        private const uint ENABLE_EXTENDED_FLAGS  = 0x0080;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetStdHandle(int nStdHandle);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetConsoleMode(
            IntPtr hConsoleHandle,
            out uint lpMode
        );

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool SetConsoleMode(
            IntPtr hConsoleHandle,
            uint dwMode
        );

        public static void DisableQuickEdit()
        {
            IntPtr handle = GetStdHandle(STD_INPUT_HANDLE);

            if (handle == IntPtr.Zero || handle == new IntPtr(-1))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            if (!GetConsoleMode(handle, out uint mode))
                throw new Win32Exception(Marshal.GetLastWin32Error());

            mode |= ENABLE_EXTENDED_FLAGS;
            mode &= ~ENABLE_QUICK_EDIT_MODE;

            if (!SetConsoleMode(handle, mode))
                throw new Win32Exception(Marshal.GetLastWin32Error());
        }
    }
}
'@
    }

    [Native.ConsoleMode]::DisableQuickEdit()
}

try {
    Disable-ConsoleQuickEdit
} catch {
    Write-Warning "Could not disable QuickEdit mode: $($_.Exception.Message)"
}