using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace SupportTuoiBeo.Data.Repositories
{
    public interface IGenericRepository<T> where T : class
    {
        Task<T> GetByIdAsync(Guid id);

        Task<IReadOnlyList<T>> GetPagedReponseAsync(int pageNumber, int pageSize);

        Task<IReadOnlyList<T>> GetAllAsync();

        Task<T> AddAsync(T entity);

        Task UpdateAsync(T entity);

        Task DeleteAsync(T entity);
    }
}
