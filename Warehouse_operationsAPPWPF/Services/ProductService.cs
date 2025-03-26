using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Warehouse_operationsAPPWPF.Models;

namespace Warehouse_operationsAPPWPF.Services
{
    internal class ProductService
    {
        private readonly HttpClient _httpClient;

        public ProductService()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/") };
        }

        public async Task<List<Product>> GetProductsByWarehouseAsync(int warehouseId)
        {
            var response = await _httpClient.GetAsync($"/api/products/byWarehouse/{warehouseId}");
            response.EnsureSuccessStatusCode();
            var json = await response.Content.ReadAsStringAsync();
            return JsonConvert.DeserializeObject<List<Product>>(json);
        }
    }
}