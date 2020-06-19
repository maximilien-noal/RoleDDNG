using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Serialization
{
    public interface IAsyncSerializer<T>
    {
        Task<T> DeserializeAsync<TModel>(string filePath)
            where TModel : T, new();

        Task SerializeAsync(string filePath, T objectToSerialize);
    }
}