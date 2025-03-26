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
    internal class ApiServiceOstatki
    {
        private readonly HttpClient _httpClient;

        public ApiServiceOstatki()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/api/Ostatki/") };
        }

        internal async Task<List<Ostatki>> GetOstatkisAsync()
        {
            var response = await _httpClient.GetAsync("");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<Ostatki>>();
        }

        internal async Task AddOstatkiAsync(Ostatki ostatki)
        {
            var response = await _httpClient.PostAsJsonAsync($"POST?id_warehouses={ostatki.id_warehouses}&id_Product={ostatki.id_Product}", ostatki);
            response.EnsureSuccessStatusCode();
        }

        internal async Task UpdateOstatkiAsync(Ostatki ostatki)
        {
            var response = await _httpClient.PutAsJsonAsync($"PUT/{ostatki.id_Ostatki}?id_warehouses={ostatki.id_warehouses}&id_Product={ostatki.id_Product}", ostatki);
            response.EnsureSuccessStatusCode();
        }

        internal async Task DeleteOstatkiAsync(int OstatkiId)
        {
            var response = await _httpClient.DeleteAsync($"DELETE/{OstatkiId}");
            response.EnsureSuccessStatusCode();
        }
    }
}
