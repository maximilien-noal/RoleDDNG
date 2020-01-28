﻿using System;
using System.Collections.Generic;

using RoleDDNG.Interfaces.Backgrounds;
using RoleDDNG.OSServices.CrossPlatform;

namespace RoleDDNG.Role.Backgrounds
{
    public class BackgroundSource : IBackgroundSource
    {
        private List<string> _backgrounds = new List<string>();

        public BackgroundSource()
        {
            if (DateTime.Now.Month == 12)
            {
                _backgrounds.Add("Assets/backgrounds/christmas.jpg");
                _backgrounds.Add("Assets/backgrounds/christmas2.png");
                _backgrounds.Add("Assets/backgrounds/christmas3.jpg");
                _backgrounds.Add("Assets/backgrounds/christmasEve.jpg");
                _backgrounds.Add("Assets/backgrounds/IceQueen.png");
                _backgrounds.Add("Assets/backgrounds/imp.png");
            }
            else if (DateTime.Now.Month == 10)
            {
                _backgrounds.Add("Assets/backgrounds/sorceress.png");
                _backgrounds.Add("Assets/backgrounds/halloween.jpg");
                _backgrounds.Add("Assets/backgrounds/halloween2.png");
                _backgrounds.Add("Assets/backgrounds/halloween3.png");
                _backgrounds.Add("Assets/backgrounds/halloween4.jpg");
                _backgrounds.Add("Assets/backgrounds/halloween5.png");
                _backgrounds.Add("Assets/backgrounds/halloween6.png");
                _backgrounds.Add("Assets/backgrounds/halloween7.png");
                _backgrounds.Add("Assets/backgrounds/halloween8.png");
            }
            else
            {
                _backgrounds.Add("Assets/backgrounds/original.png");
                _backgrounds.Add("Assets/backgrounds/citiesInTheSky.png");
                _backgrounds.Add("Assets/backgrounds/deer.jpg");
                _backgrounds.Add("Assets/backgrounds/deer2.jpg");
                _backgrounds.Add("Assets/backgrounds/diceTable.jpg");
                _backgrounds.Add("Assets/backgrounds/dwarfHouse.jpg");
                _backgrounds.Add("Assets/backgrounds/faery.png");
                _backgrounds.Add("Assets/backgrounds/forest.png");
                _backgrounds.Add("Assets/backgrounds/forest2.jpg");
                _backgrounds.Add("Assets/backgrounds/LotsOfDices.png");
                _backgrounds.Add("Assets/backgrounds/merchant.jpg");
                _backgrounds.Add("Assets/backgrounds/original.png");
                _backgrounds.Add("Assets/backgrounds/sword.png");
                _backgrounds.Add("Assets/backgrounds/unicorn.png");
                _backgrounds.Add("Assets/backgrounds/warrior.jpg");
            }
        }

        public string GetBackgroundSource()
        {
            return _backgrounds[new StaticRNG().GetStandardRNG().Next(0, _backgrounds.Count)];
        }
    }
}