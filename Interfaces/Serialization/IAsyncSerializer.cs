using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Serialization
{
    public interface IAsyncSerializer<T>
    {
        public Task<T> DeserializeAsync(string filePath);

        public Task SerializeAsync(string filePath, T objectToSerialize);
    }
}