using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Threading.Tasks;

namespace RoleDDNG.Interfaces.Repository
{
    public interface IAsyncRepository<T> where T : BaseEntity
    {
        Task AddAsync(T entity);

        Task<int> CountAllAsync();

        Task<int> CountWhereAsync(Expression<Func<T, bool>> predicate);

        Task<T> FirstOrDefaultAsync(Expression<Func<T, bool>> predicate);

        Task<IEnumerable<T>> GetAllAsync();

        Task<T> GetByIdAsync(int id);

        Task<IEnumerable<T>> GetWhereAsync(Expression<Func<T, bool>> predicate);

        Task RemoveAsync(T entity);

        Task UpdateAsync(T entity);
    }
}