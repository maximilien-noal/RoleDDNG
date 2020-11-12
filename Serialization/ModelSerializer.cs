using RoleDDNG.Interfaces.Serialization;

using System.IO;
using System.Text.Json;
using System.Threading.Tasks;

namespace RoleDDNG.Serialization
{
    public class ModelSerializer<T> : IAsyncSerializer<T>
    {
        public async Task<T> DeserializeAsync<TModel>(string filePath)
            where TModel : T, new()
        {
            try
            {
                using var reader = File.OpenRead(filePath);
                TModel? t = await JsonSerializer.DeserializeAsync<TModel>(reader).ConfigureAwait(false);
                if (t is null)
                {
                    return new TModel();
                }
                return t;
            }
            catch (JsonException)
            {
                return await new ValueTask<TModel>(new TModel()).ConfigureAwait(false);
            }
        }

        public async Task SerializeAsync(string filePath, T objectToSerialize)
        {
            using var fileStream = File.Create(filePath);
            await JsonSerializer.SerializeAsync(fileStream, objectToSerialize).ConfigureAwait(false);
        }
    }
}