using RoleDDNG.Models.Structs;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace RoleDDNG.PInvoke
{
    public static class NativeMethods
    {
        [DllImport("user32.dll")]
        private static extern bool SetWindowPlacement(IntPtr hWnd, [In] ref WindowPlacement lpwndpl);

        [DllImport("user32.dll")]
        private static extern bool GetWindowPlacement(IntPtr hWnd, out WindowPlacement lpwndpl);
    }
}