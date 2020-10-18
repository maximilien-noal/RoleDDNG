using System;

namespace RoleDDNG.Models
{
    internal static class OleDateTime
    {
        public static DateTime WithoutMs(this DateTime dt)
        {
            return dt.AddTicks(-(dt.Ticks % TimeSpan.TicksPerSecond));
        }
    }
}