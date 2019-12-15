using System;
using System.Threading.Tasks;

namespace Interfaces
{
    public interface IAsyncLoadSave
    {
        public Task SaveModel<T>(T model) where T : Type;

        public Task<T> LoadModel<T>() where T : new();
    }
}