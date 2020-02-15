using System.Threading.Tasks;
using System.Windows.Threading;

using Microsoft.Win32;

using RoleDDNG.Interfaces.Dialogs;

namespace RoleDDNG.OSServices.Windows.Dialogs
{
    public class Win32FileDialog : IFileDialog
    {
        public async Task<string> TryOpenUserChosenFileAsync(string dialogTitle, string defaultExtension)
        {
            var filename = "";
            await Dispatcher.CurrentDispatcher.InvokeAsync(() =>
            {
                var dialog = new OpenFileDialog
                {
                    CheckFileExists = true,
                    Multiselect = false,
                    DefaultExt = defaultExtension,
                    Filter = $"{defaultExtension.Substring(1)} ({defaultExtension})|*{defaultExtension}",
                    Title = dialogTitle
                };
                if (dialog.ShowDialog() == true)
                {
                    filename = dialog.FileName;
                }
            });
            return filename;
        }
    }
}