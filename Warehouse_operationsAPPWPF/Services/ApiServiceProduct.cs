using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Warehouse_operationsAPPWPF.Models;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using System.Text.Json;

namespace Warehouse_operationsAPPWPF.Services
{
    public class ApiServiceProduct
    {
        private readonly HttpClient _httpClient;

        public ApiServiceProduct()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("http://localhost:5093/api/Product/") };
        }

        internal async Task<List<Product>> GetProductsAsync()
        {
            var response = await _httpClient.GetAsync("");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<Product>>();
        }

        internal async Task AddProductAsync(Product product)
        {
            var response = await _httpClient.PostAsJsonAsync("POST", product); // Теперь отправляем JSON полностью
            response.EnsureSuccessStatusCode();
        }

        internal async Task UpdateProductAsync(Product product)
        {
            var response = await _httpClient.PutAsJsonAsync($"PUT/{product.id_Product}", product);
            response.EnsureSuccessStatusCode();
        }

        internal async Task DeleteProductAsync(int productId)
        {
            var response = await _httpClient.DeleteAsync($"DELETE/{productId}");
            response.EnsureSuccessStatusCode();
        }
    }
}