using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Users
    {
        public int id_users { get; set; }
        public string FIO { get; set; }
        public string Login { get; set; }
        public string password { get; set; }
        public int id_doljnosti { get; set; }
    }
}
