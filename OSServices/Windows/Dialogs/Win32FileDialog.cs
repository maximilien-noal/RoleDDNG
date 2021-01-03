using Microsoft.Win32;

using RoleDDNG.Interfaces.Dialogs;

using System.Collections.Generic;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace RoleDDNG.OSServices.Windows.Dialogs
{
    public class Win32FileDialog : IFileDialog
    {
        public async Task<string> OpenFileDialogAsync(string dialogTitle, string defaultExtension)
        {
            var filename = "";
            await Dispatcher.CurrentDispatcher.InvokeAsync(() =>
            {
                var dialog = new OpenFileDialog
                {
                    CheckFileExists = true,
                    Multiselect = false,
                    DefaultExt = defaultExtension,
                    Filter = $"{defaultExtension[1..]} ({defaultExtension})|*{defaultExtension}",
                    Title = dialogTitle
                };
                if (dialog.ShowDialog(Application.Current.MainWindow) == true)
                {
                    filename = dialog.FileName;
                }
            });
            return filename;
        }

        public async Task<IEnumerable<string>> OpenFilesDialogAsync(string dialogTitle, string defaultExtension)
        {
            IEnumerable<string>? filenames = null;
            await Dispatcher.CurrentDispatcher.InvokeAsync(() =>
            {
                var dialog = new OpenFileDialog
                {
                    CheckFileExists = true,
                    Multiselect = true,
                    DefaultExt = defaultExtension,
                    Filter = $"{defaultExtension[1..]} ({defaultExtension})|*{defaultExtension}",
                    Title = dialogTitle
                };
                if (dialog.ShowDialog(Application.Current.MainWindow) == true)
                {
                    filenames = dialog.FileNames;
                }
            });
            return filenames is null ? new List<string>() : filenames;
        }
    }
}