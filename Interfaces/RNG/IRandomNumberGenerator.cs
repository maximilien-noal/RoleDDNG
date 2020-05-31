﻿using System;

namespace RoleDDNG.Interfaces.RNG
{
    public interface IRandomNumberGenerator
    {
        Random GetLimitedRNG();

        Random GetStandardRNG();
    }
}