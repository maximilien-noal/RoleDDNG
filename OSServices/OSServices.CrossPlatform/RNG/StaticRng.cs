using System;

using RoleDDNG.Interfaces.RNG;

namespace RoleDDNG.OSServices.CrossPlatform
{
    public class StaticRNG : IRandomNumberGenerator
    {
        private static readonly Random _limitedRNG = new Random(1000);

        private static readonly Random _standardRNG = new Random();

        public Random GetLimitedRNG()
        {
            return _limitedRNG;
        }

        public Random GetStandardRNG()
        {
            return _standardRNG;
        }
    }
}