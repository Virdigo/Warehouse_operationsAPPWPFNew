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
    internal class ApiServiceInformation
    {
        private readonly HttpClient _httpClient;

        public ApiServiceInformation()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/api/Information_about_documents/") };
        }

        internal async Task<List<Information_about_documents>> GetInformationAsync()
        {
            var response = await _httpClient.GetAsync("");
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadFromJsonAsync<List<Information_about_documents>>();
        }

        internal async Task AddInformationAsync(Information_about_documents information)
        {
            var response = await _httpClient.PostAsJsonAsync($"POST?id_Product={information.id_Product}&id_doc={information.id_doc}&id_suppliers={information.id_suppliers}", information);
            response.EnsureSuccessStatusCode();
        }

        internal async Task UpdateInformationAsync(Information_about_documents information)
        {
            var response = await _httpClient.PutAsJsonAsync($"PUT/{information.id_inf_doc}?id_Product={information.id_Product}&id_doc={information.id_doc}&id_suppliers={information.id_suppliers}", information);
            response.EnsureSuccessStatusCode();
        }

        internal async Task DeleteInformationAsync(int id_inf_doc)
        {
            var response = await _httpClient.DeleteAsync($"DELETE/{id_inf_doc}");
            response.EnsureSuccessStatusCode();
        }
    }
}