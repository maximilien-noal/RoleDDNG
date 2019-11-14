using System;

namespace RoleDDNG.ViewModels.Interfaces
{
    public interface IContent
    {
        string Title { get; }
        bool CanClose { get; }

        event EventHandler Closing;
    }
}