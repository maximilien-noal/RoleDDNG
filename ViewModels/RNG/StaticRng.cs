using System;

namespace RoleDDNG.ViewModels.RNG
{
    internal static class StaticRNG
    {
        public static Random RNG { get; private set; } = new Random();

        public static Random LimitedRNG { get; private set; } = new Random(1000);
    }
}