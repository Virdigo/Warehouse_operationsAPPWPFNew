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
    internal class ApiServiceLogin
    {
        private readonly HttpClient _httpClient;

        public ApiServiceLogin()
        {
            _httpClient = new HttpClient { BaseAddress = new Uri("https://localhost:7034/api/") };
        }

        internal async Task<Users> AuthenticateAsync(string login, string password)
        {
            var loginDto = new LoginDto { Login = login, Password = password };
            var response = await _httpClient.PostAsJsonAsync("Users/authenticate", loginDto);
            if (response.IsSuccessStatusCode)
            {
                return await response.Content.ReadFromJsonAsync<Users>();
            }
            return null;
        }
    }
}