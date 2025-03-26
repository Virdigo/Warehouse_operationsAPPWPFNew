using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Product
    {
        public int id_Product { get; set; }
        public string Name { get; set; }
        public string vendor_code { get; set; }
        public int Price { get; set; }
        public int id_product_type { get; set; } // Нужно для API
        public int id_unit { get; set; }         // Нужно для API   
        public string ProductTypeName { get; set; } // Вместо id_product_type
        public string UnitName { get; set; }       // Вместо id_unit
    }
}
