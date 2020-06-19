using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Dialogs
{
    public interface IFileDialog
    {
        Task<string> OpenFileDialogAsync(string dialogTitle, string defaultExtension);
    }
}