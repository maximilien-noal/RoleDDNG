using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Serialization
{
    public interface ISerializer<T>
    {
        public T Deserialize(string filePath);

        public void Serialize(string filePath, T objectToSerialize);
    }
}