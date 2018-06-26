using System;
using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TesteEntity
{
    public class Program
    {        
        static void Main(string[] args)
        {
            var dbs = DbNames();
            int soma = 0;
            string connString = @"Server=DESKTOP-5K9ILHM\SQLEXPRESS;Database={0};User Id=sa; Password=sa@server;";
            for (int i = 0; i < dbs.Count; i++)
            { 
                using(var ctx = new Context(String.Format(connString, dbs[i])))
                {
                    soma += ctx.Usuarios.Sum(s => s.Quantidade);
                }

            }
            Console.WriteLine("Soma: " + soma);
            Console.ReadKey();
        }

        public static List<string> DbNames()
        {
            return new List<string>()
            {
                "BancoA",
                "BancoB",
                "BancoC",
                "BancoD"
            };
        }
    }
}
