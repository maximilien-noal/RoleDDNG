using System;
using System.Runtime.InteropServices;

using RoleDDNG.Models.Structs;

namespace RoleDDNG.OSServices.Windows.PInvoke
{
    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl);
    }
}