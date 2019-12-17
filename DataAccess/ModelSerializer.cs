using RoleDDNG.Interfaces.Serialization;

using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;

namespace DataAccess
{
    public class ModelSerializer<T> : IAsyncSerializer<T>
    {
        public Task<T> DeserializeAsync(string filePath)
        {
            return Task.Run<T>(() =>
            {
                XmlReaderSettings settings = new XmlReaderSettings() { XmlResolver = null };
                using XmlReader reader = XmlReader.Create(filePath, settings);
                var serializer = new XmlSerializer(typeof(T));
                return (T)serializer.Deserialize(reader);
            });
        }

        public Task SerializeAsync(string filePath, T objectToSerialize)
        {
            return Task.Run(() =>
            {
                XmlSerializer serializer = new XmlSerializer(typeof(T));
                using var writer = new StreamWriter(filePath, false, Encoding.Unicode);
                serializer.Serialize(writer, objectToSerialize);
                writer.Close();
            });
        }
    }
}