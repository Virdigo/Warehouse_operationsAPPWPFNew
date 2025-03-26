using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Warehouse_operationsAPPWPF.Pages.Accountant;
using Warehouse_operationsAPPWPF.Pages.WorkWarehouse;
using Warehouse_operationsAPPWPF.Services;


namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для LoginPage.xaml
    /// </summary>
    public partial class LoginPage : Page
    {
        private readonly ApiServiceLogin _apiService;

        public LoginPage()
        {
            InitializeComponent();
            _apiService = new ApiServiceLogin();
        }

        private async void LoginButton_Click(object sender, RoutedEventArgs e)
        {
            var login = LoginTextBox.Text;
            var password = PasswordBox.Password;

            try
            {
                var user = await _apiService.AuthenticateAsync(login, password);
                if (user != null)
                {
                    switch (user.id_doljnosti)
                    {
                        case 1:
                            NavigationService.Navigate(new PageMenu());
                            break;
                        case 2:
                            NavigationService.Navigate(new ProductsPageWarehouse());
                            break;
                        case 3:
                            NavigationService.Navigate(new OstatkiPageAccountant());
                            break;
                        default:
                            MessageBox.Show("Неизвестная должность.");
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Invalid login or password");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}");
            }
        }

    }
    }
