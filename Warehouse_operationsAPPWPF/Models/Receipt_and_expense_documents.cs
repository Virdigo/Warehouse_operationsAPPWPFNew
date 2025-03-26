using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Warehouse_operationsAPPWPF.Models
{
    internal class Receipt_and_expense_documents
    {
        public int id_doc { get; set; }
        public DateTime date { get; set; }
        public bool ReceiptAndexpense_documents { get; set; }
        public int id_users { get; set; }
        public string UsersName { get; set; }
    }
}
