using System;
using System.Collections.Generic;
using System.Text;

namespace RoleDDNG.ViewModels.RNG
{
    internal static class StaticRNG
    {
        public static Random RNG { get; private set; } = new Random();
    }
}