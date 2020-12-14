using SupportTuoiBeo.Data.Contexts;
using SupportTuoiBeo.Data.Entities;
using System;
using System.Collections.Generic;
using System.Text;

namespace SupportTuoiBeo.Data.Repositories
{
    public class UserDetailsRepository : GenericRepository<UserDetails>, IUserDetailsRepository
    {
        public UserDetailsRepository(ApplicationDbContext dbContext) : base(dbContext)
        {
        }
    }
}
