using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Dialogs
{
    public interface IFileDialog
    {
        Task<string> TryOpenUserChosenFileAsync(string dialogTitle, string defaultExtension);
    }
}