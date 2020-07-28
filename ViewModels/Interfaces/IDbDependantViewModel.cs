using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Interfaces
{
    public interface IDbDependantViewModel
    {
        public Task LoadDbDataAsync();
    }
}