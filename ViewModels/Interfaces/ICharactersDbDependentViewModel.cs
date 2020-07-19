using System.Threading.Tasks;

namespace RoleDDNG.ViewModels.Interfaces
{
    public interface ICharactersDbDependentViewModel
    {
        public Task LoadCharactersDbDataAsync();
    }
}