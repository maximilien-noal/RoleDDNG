using System;
using System.Collections.Generic;
using System.Text;
using GalaSoft.MvvmLight;

namespace RoleDDNG.ViewModels.Interfaces
{
    public interface IBusyStateNotifier
    {
        public bool IsBusy { get; set; }
    }
}