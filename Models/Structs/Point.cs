using System;
using System.Runtime.InteropServices;

namespace RoleDDNG.Models.Structs
{
    [Serializable]
    [StructLayout(LayoutKind.Sequential)]
    public struct Point : IEquatable<Point>
    {
        public int X { get; set; }

        public int Y { get; set; }

        public Point(int x, int y)
        {
            X = x;
            Y = y;
        }

#pragma warning disable CS8765 // Nullability of type of parameter doesn't match overridden member (possibly because of nullability attributes).

        public override bool Equals(object obj)
#pragma warning restore CS8765 // Nullability of type of parameter doesn't match overridden member (possibly because of nullability attributes).
        {
            if (obj is null) { return false; }
            if (!(obj is Point)) { return false; }
            return Equals((Point)obj);
        }

        public bool Equals(Point other)
        {
            return other.X == X && other.Y == Y;
        }

        public override int GetHashCode()
        {
            return HashCode.Combine(this.X, this.Y, base.GetHashCode());
        }

        public static bool operator ==(Point left, Point right)
        {
            return left.Equals(right);
        }

        public static bool operator !=(Point left, Point right)
        {
            return !(left == right);
        }
    }
}