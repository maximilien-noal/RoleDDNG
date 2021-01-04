using System;

namespace RoleDDNG.Interfaces.Taskbar
{
    public interface ITaskbarProgress
    {
        void SetState(IntPtr windowHandle, TaskbarStates taskbarState);

        void SetValue(IntPtr windowHandle, double progressValue, double progressMax);
    }
}