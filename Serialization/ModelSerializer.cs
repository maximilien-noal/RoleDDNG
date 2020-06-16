using System.IO;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

using RoleDDNG.Interfaces.Serialization;

namespace RoleDDNG.Serialization
{
    public class ModelSerializer<T> : IAsyncSerializer<T>
    {
        public async Task<T> DeserializeAsync(string filePath)
        {
            using var reader = File.OpenRead(filePath);
            return await JsonSerializer.DeserializeAsync<T>(reader).ConfigureAwait(false);
        }

        public async Task SerializeAsync(string filePath, T objectToSerialize)
        {
            using var writer = File.OpenWrite(filePath);
            await JsonSerializer.SerializeAsync<T>(writer, objectToSerialize).ConfigureAwait(false);
        }
    }
}