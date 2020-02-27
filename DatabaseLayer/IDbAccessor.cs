using System.Collections.Generic;
using System.Threading.Tasks;

namespace RoleDDNG.DatabaseLayer
{
    public interface IDbAccessor
    {
        Task<IEnumerable<T>> GetQueryDataAsync<T>(string sqlQuery);
    }
}