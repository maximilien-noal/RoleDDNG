using System;
using System.Runtime.InteropServices;

using RoleDDNG.Interfaces.Window;
using RoleDDNG.Models.Structs;
using RoleDDNG.OSServices.Windows.PInvoke;

namespace RoleDDNG.OSServices.Windows.Window
{
    public class HwndPlacer : IWindowPlacer
    {
        public bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl)
        {
            return NativeMethods.GetWindowPlacement(hWnd, out lpwndpl);
        }

        public bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl)
        {
            return NativeMethods.SetWindowPlacement(hWnd, ref lpwndpl);
        }
    }
}