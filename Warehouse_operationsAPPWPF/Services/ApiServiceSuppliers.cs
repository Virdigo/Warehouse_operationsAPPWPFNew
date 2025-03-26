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
    internal class ApiServiceSuppliers
    {
        private readonly HttpClient _httpClient;

        public ApiServiceSuppliers()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/api/Suppliers/") };
        }

        internal async Task<List<Suppliers>> GetSuppliersAsync()
        {
            var response = await _httpClient.GetAsync("");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<Suppliers>>();
        }

        internal async Task AddSuppliersAsync(Suppliers suppliers)
        {
            var response = await _httpClient.PostAsJsonAsync("POST", suppliers);
            response.EnsureSuccessStatusCode();
        }

        internal async Task UpdateSuppliersAsync(Suppliers suppliers)
        {
            var response = await _httpClient.PutAsJsonAsync($"PUT/{suppliers.id_suppliers}", suppliers);
            response.EnsureSuccessStatusCode();
        }

        internal async Task DeleteSuppliersAsync(int suppliersId)
        {
            var response = await _httpClient.DeleteAsync($"DELETE/{suppliersId}");
            response.EnsureSuccessStatusCode();
        }
    }
}