using System;

using RoleDDNG.Interfaces.RNG;

namespace RoleDDNG.OSServices.CrossPlatform
{
    public class StaticRng : IRandomNumberGenerator
    {
        private static readonly Random _standardRNG = new Random();

        public Random GetRng()
        {
            return _standardRNG;
        }
    }
}