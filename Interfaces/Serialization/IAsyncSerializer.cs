using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Serialization
{
    public interface IAsyncSerializer<T>
    {
        Task<T> DeserializeAsync(string filePath);

        Task SerializeAsync(string filePath, T objectToSerialize);
    }
}