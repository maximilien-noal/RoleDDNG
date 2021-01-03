using System.Collections.Generic;
using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Dialogs
{
    public interface IFileDialog
    {
        Task<string> OpenFileDialogAsync(string dialogTitle, string defaultExtension);

        Task<IEnumerable<string>> OpenFilesDialogAsync(string dialogTitle, string defaultExtension);
    }
}