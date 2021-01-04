using RoleDDNG.Interfaces.Taskbar;

using System;
using System.Runtime.InteropServices;

namespace RoleDDNG.OSServices.Windows.Taskbar
{
    public class TaskbarProgress : ITaskbarProgress
    {
        [ComImport()]
        [Guid("ea1afb91-9e28-4b86-90e9-9e9f8a5eefaf")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface ITaskbarList3
        {
            // ITaskbarList
            [PreserveSig]
            void HrInit();

            [PreserveSig]
            void AddTab(IntPtr hwnd);

            [PreserveSig]
            void DeleteTab(IntPtr hwnd);

            [PreserveSig]
            void ActivateTab(IntPtr hwnd);

            [PreserveSig]
            void SetActiveAlt(IntPtr hwnd);

            // ITaskbarList2
            [PreserveSig]
            void MarkFullscreenWindow(IntPtr hwnd, [MarshalAs(UnmanagedType.Bool)] bool fFullscreen);

            // ITaskbarList3
            [PreserveSig]
            void SetProgressValue(IntPtr hwnd, UInt64 ullCompleted, UInt64 ullTotal);

            [PreserveSig]
            void SetProgressState(IntPtr hwnd, TaskbarStates state);
        }

        [ComImport()]
        [Guid("56fdf344-fd6d-11d0-958a-006097c9a090")]
        [ClassInterface(ClassInterfaceType.None)]
        private class TaskbarInstance
        {
        }

        private static readonly ITaskbarList3 _taskbarInstance = (ITaskbarList3)new TaskbarInstance();
        private static readonly bool _taskbarSupported = Environment.OSVersion.Version >= new Version(6, 1);

        public void SetState(IntPtr windowHandle, TaskbarStates taskbarState)
        {
            if (_taskbarSupported)
            {
                _taskbarInstance.SetProgressState(windowHandle, taskbarState);
            }
        }

        public void SetValue(IntPtr windowHandle, double progressValue, double progressMax)
        {
            if (_taskbarSupported)
            {
                _taskbarInstance.SetProgressValue(windowHandle, (ulong)progressValue, (ulong)progressMax);
            }
        }
    }
}