using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

using RoleDDNG.Interfaces.Serialization;

namespace RoleDDNG.Serialization
{
    public class ModelSerializer<T> : IAsyncSerializer<T>
    {
        public async Task<T> DeserializeAsync(string filePath)
        {
            var deserializedMdodel = await Task.Run(() =>
            {
                var settings = new XmlReaderSettings { XmlResolver = null };
                using var reader = XmlReader.Create(filePath, settings);
                var serializer = new XmlSerializer(typeof(T));
                return (T)serializer.Deserialize(reader);
            }).ConfigureAwait(false);
            return deserializedMdodel;
        }

        public async Task SerializeAsync(string filePath, T objectToSerialize)
        {
            await Task.Run(() =>
            {
                var serializer = new XmlSerializer(typeof(T));
                using var writer = new StreamWriter(filePath, false, Encoding.Unicode);
                serializer.Serialize(writer, objectToSerialize);
                writer.Close();
            }).ConfigureAwait(false);
        }
    }
}