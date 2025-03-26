using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Warehouses
    {
        public int id_warehouses { get; set; }
        public string Name { get; set; }
        public string address { get; set; }
        public int id_users { get; set; }
        public string UsersName { get; set; }
    }
}
