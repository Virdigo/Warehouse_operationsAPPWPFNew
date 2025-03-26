using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Information_about_documents
    {
        public int id_inf_doc { get; set; }
        public int id_Product { get; set; }
        public int Quanity { get; set; }
        public int id_doc { get; set; }
        public int id_suppliers { get; set; }
        public int Cost { get; set; }
        public int Price { get; set; }
        public string ProductName { get; set; }
        public string Receipt_and_expense_documentsName { get; set; }
        public string SuppliersName { get; set; }
    }
}
