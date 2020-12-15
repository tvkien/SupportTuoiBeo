using Microsoft.EntityFrameworkCore;
using SupportTuoiBeo.Data.Entities;

namespace SupportTuoiBeo.Data.Contexts
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext()
        {

        }
        public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options) { }

        public DbSet<UserDetails> UserDetails { get; set; }

        protected override void OnConfiguring(DbContextOptionsBuilder builder)
        {
            builder.UseSqlServer("Server=.;Database=SupportTuoiBeo;Trusted_Connection=True;");
        }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<UserDetails>().ToTable("UserDetails").HasKey(x => x.Id);
            modelBuilder.Entity<UserDetails>().Property(x => x.Id).UseIdentityColumn();
            //modelBuilder.Entity<UserDetails>().HasIndex(x => x.MaKH).IsUnique();
        }
    }
}