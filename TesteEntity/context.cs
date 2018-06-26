using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesteEntity
{
    public class Context : DbContext
    {
        public Context(string conn) : base(conn)
        {

        }

        public DbSet<Usuarios> Usuarios { get; set; }
    }
}
