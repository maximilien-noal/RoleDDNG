using System.Collections.Generic;
using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Repository
{
    public interface IAsyncRepository<T>
    {
        //Task AddAsync(T entity);

        //Task<int> CountAllAsync();

        //Task<int> CountWhereAsync(Expression<Func<T, bool>> predicate);

        //Task<T> FirstOrDefaultAsync(Expression<Func<T, bool>> predicate);

        Task<IEnumerable<T>> GetAllAsync(string databasePath, string tableName);

        //Task<T> GetByIdAsync(int id);

        //Task<IEnumerable<T>> GetWhereAsync(Expression<Func<T, bool>> predicate);

        //Task RemoveAsync(T entity);

        //Task UpdateAsync(T entity);
    }
}