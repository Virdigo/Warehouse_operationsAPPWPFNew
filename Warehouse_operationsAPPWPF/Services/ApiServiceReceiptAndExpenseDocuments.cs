using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Json;
using System.Text;
using System.Threading.Tasks;
using Warehouse_operationsAPPWPF.Models;

namespace Warehouse_operationsAPPWPF.Services
{
    public class ApiServiceReceiptAndExpenseDocuments
    {
        private readonly HttpClient _httpClient;

        public ApiServiceReceiptAndExpenseDocuments()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/api/Receipt_and_expense_documents/") };
        }

        internal async Task<List<Receipt_and_expense_documents>> GetReceiptAndExpenseDocumentsList()
        {
            var response = await _httpClient.GetAsync("");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<Receipt_and_expense_documents>>();
        }

        internal async Task AddReceipt_and_expense_documentsAsync(Receipt_and_expense_documents receipt_and_expense_documents)
        {
            var response = await _httpClient.PostAsJsonAsync($"POST?id_users={receipt_and_expense_documents.id_users}", receipt_and_expense_documents);
            response.EnsureSuccessStatusCode();
        }

        internal async Task UpdateReceipt_and_expense_documentsAsync(Receipt_and_expense_documents receipt_and_expense_documents)
        {
            var response = await _httpClient.PutAsJsonAsync($"PUT/{receipt_and_expense_documents.id_doc}?id_users={receipt_and_expense_documents.id_users}", receipt_and_expense_documents);
            response.EnsureSuccessStatusCode();
        }

        internal async Task DeleteReceipt_and_expense_documentsAsync(int docId)
        {
            var response = await _httpClient.DeleteAsync($"DELETE/{docId}");
            response.EnsureSuccessStatusCode();
        }
    }
}
