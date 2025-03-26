using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Ostatki
    {
        public int id_Ostatki { get; set; }
        public int id_warehouses { get; set; }
        public int id_Product { get; set; }
        public int Quantity_Ostatki { get; set; }
        public string WarehousesName { get; set; }
        public string ProductName { get; set; }
    }
}
