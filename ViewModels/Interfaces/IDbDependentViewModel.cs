using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Interfaces
{
    public interface IDbDependentViewModel
    {
        public Task LoadDbDataAsync();
    }
}