using System;
using System.Runtime.InteropServices;

namespace RoleDDNG.Models.Structs
{
    [Serializable]
    [StructLayout(LayoutKind.Sequential)]
    public struct Rect : IEquatable<Rect>
    {
        public int Left { get; set; }
        public int Top { get; set; }
        public int Right { get; set; }
        public int Bottom { get; set; }

        public Rect(int left, int top, int right, int bottom)
        {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        public bool Equals(Rect other)
        {
            return other.Left == Left && other.Top == Top && other.Right == Right && other.Bottom == Bottom;
        }

        public override bool Equals(object obj)
        {
            if (obj is null) { return false; }
            if (obj is Rect == false) { return false; }
            return Equals((Rect)obj);
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(this, base.GetHashCode());
        }

        public static bool operator ==(Rect left, Rect right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(Rect left, Rect right)
        {
            return !(left == right);
        }
    }
}