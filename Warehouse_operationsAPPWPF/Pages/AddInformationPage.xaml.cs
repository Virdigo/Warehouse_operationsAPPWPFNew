using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using Warehouse_operationsAPPWPF.Models;
using Warehouse_operationsAPPWPF.Services;

namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для AddInformationPage.xaml
    /// </summary>
    public partial class AddInformationPage : Page
    {
        private readonly ApiServiceInformation _apiService;
        private readonly Information_about_documents _information;


        internal AddInformationPage(Information_about_documents information = null)
        {
            InitializeComponent();
            _apiService = new ApiServiceInformation();
            _information = information ?? new Information_about_documents();

                ProductIdTextBox.Text = _information.id_Product.ToString();
                QuantityTextBox.Text = _information.Quanity.ToString();
                DocumentIdTextBox.Text = _information.id_doc.ToString();
                SupplierIdTextBox.Text = _information.id_suppliers.ToString();
                CostTextBox.Text = _information.Cost.ToString();
                PriceTextBox.Text = _information.Price.ToString();
            
        }
        private void FilterTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только ввод цифр и знак минус
            e.Handled = !Regex.IsMatch(e.Text, @"^[0-9]+$");
        }


        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _information.id_Product = int.Parse(ProductIdTextBox.Text);
                _information.Quanity = int.Parse(QuantityTextBox.Text);
                _information.id_doc = int.Parse(DocumentIdTextBox.Text);
                _information.id_suppliers = int.Parse(SupplierIdTextBox.Text);
                _information.Cost = int.Parse(CostTextBox.Text);
                _information.Price = int.Parse(PriceTextBox.Text);

                if (_information.id_inf_doc == 0)
                {
                    await _apiService.AddInformationAsync(_information);
                    MessageBox.Show("Информация успешно добавлена");
                }
                else
                {
                    await _apiService.UpdateInformationAsync(_information);
                    MessageBox.Show("Информация успешно обновлена");
                }

                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении информации: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}
