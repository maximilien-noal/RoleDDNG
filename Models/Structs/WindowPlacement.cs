using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace RoleDDNG.Models.Structs
{
    /// <summary>
    /// Stores the position, size, and state of a window
    /// </summary>
    [Serializable]
    [StructLayout(LayoutKind.Sequential)]
    public struct WindowPlacement : IEquatable<WindowPlacement>
    {
        public int Length { get; set; }
        public int Flags { get; set; }
        public int ShowCmd { get; set; }
        public Point MinPosition { get; set; }
        public Point MaxPosition { get; set; }
        public Rect NormalPosition { get; set; }

        /// <inheritdoc />
        public bool Equals(WindowPlacement wp)
        {
            return wp.Flags == Flags && wp.Length == Length && wp.ShowCmd == ShowCmd &&
                    wp.MinPosition == MinPosition && wp.MaxPosition == MaxPosition &&
                    wp.NormalPosition.Left == NormalPosition.Left && wp.NormalPosition.Right == NormalPosition.Right &&
                    wp.NormalPosition.Top == NormalPosition.Top && wp.NormalPosition.Bottom == NormalPosition.Bottom;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(this, base.GetHashCode());
        }

        public static bool operator ==(WindowPlacement left, WindowPlacement right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(WindowPlacement left, WindowPlacement right)
        {
            return !(left == right);
        }

        public override bool Equals(object obj)
        {
            if (obj is WindowPlacement == false)
            {
                return false;
            }

            return Equals((WindowPlacement)obj);
        }
    }
}