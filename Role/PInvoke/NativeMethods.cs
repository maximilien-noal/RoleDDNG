using RoleDDNG.Models.Structs;
using System;
using System.Runtime.InteropServices;

namespace RoleDDNG.Role.PInvoke
{
    internal static class NativeMethods
    {
        [DllImport("user32.dll")]
        public static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl);

        [DllImport("user32.dll")]
        public static extern bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl);
    }
}