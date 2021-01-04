namespace RoleDDNG.Models.Sessions
{
    using System.Collections.ObjectModel;

    public class Session
    {
        public string Id { get; set; } = "";

        public ObservableCollection<object> ViewModels { get; set; } = new ObservableCollection<object>();
    }
}