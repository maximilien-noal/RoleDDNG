using System;
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
        public bool Equals(WindowPlacement other)
        {
            return other.Flags == Flags && other.Length == Length && other.ShowCmd == ShowCmd &&
                    other.MinPosition == MinPosition && other.MaxPosition == MaxPosition &&
                    other.NormalPosition.Left == NormalPosition.Left && other.NormalPosition.Right == NormalPosition.Right &&
                    other.NormalPosition.Top == NormalPosition.Top && other.NormalPosition.Bottom == NormalPosition.Bottom;
        }

        public override bool Equals(object obj)
        {
            if (!(obj is WindowPlacement))
            {
                return false;
            }

            return Equals((WindowPlacement)obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(this.Flags, this.ShowCmd, this.Length, this.NormalPosition, base.GetHashCode());
        }

        public static bool operator ==(WindowPlacement left, WindowPlacement right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(WindowPlacement left, WindowPlacement right)
        {
            return !(left == right);
        }
    }
}