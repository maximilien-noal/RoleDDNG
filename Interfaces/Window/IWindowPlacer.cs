using System;
using System.Runtime.InteropServices;

using RoleDDNG.Models.Structs;

namespace RoleDDNG.Interfaces.Window
{
    public interface IWindowPlacer
    {
        bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl);

        bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl);
    }
}