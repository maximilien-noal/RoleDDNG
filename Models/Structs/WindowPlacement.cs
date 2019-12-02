using PInvoke;

using System.Drawing;
using System.Runtime.InteropServices;

namespace RoleDDNG.Models.Structs
{
    /// <summary>
    /// Stores the position, size, and state of a window
    /// </summary>
    [StructLayout(LayoutKind.Sequential)]
    public struct WindowPlacement : System.IEquatable<WindowPlacement>
    {
        public int Length { get; set; }
        public int Flags { get; set; }
        public int ShowCmd { get; set; }
        public Point MinPosition { get; set; }
        public Point MaxPosition { get; set; }
        public RECT NormalPosition { get; set; }

        /// <inheritdoc />
        public bool Equals(WindowPlacement wp)
        {
            return wp.Flags == Flags && wp.Length == Length && wp.ShowCmd == ShowCmd &&
                    wp.MinPosition == MinPosition && wp.MaxPosition == MaxPosition &&
                    wp.NormalPosition.left == NormalPosition.left && wp.NormalPosition.right == NormalPosition.right &&
                    wp.NormalPosition.top == NormalPosition.top && wp.NormalPosition.bottom == NormalPosition.bottom;
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
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