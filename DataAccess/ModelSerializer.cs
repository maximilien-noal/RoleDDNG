using RoleDDNG.Interfaces.Serialization;

using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace DataAccess
{
    public class ModelSerializer<T> : ISerializer<T>
    {
        public T Deserialize(string filePath)
        {
            XmlReaderSettings settings = new XmlReaderSettings() { XmlResolver = null };
            using XmlReader reader = XmlReader.Create(filePath, settings);
            var serializer = new XmlSerializer(typeof(T));
            return (T)serializer.Deserialize(reader);
        }

        public void Serialize(string filePath, T objectToSerialize)
        {
            XmlSerializer serializer = new XmlSerializer(typeof(T));
            using var writer = new StreamWriter(filePath, false, Encoding.Unicode);
            serializer.Serialize(writer, objectToSerialize);
            writer.Close();
        }
    }
}