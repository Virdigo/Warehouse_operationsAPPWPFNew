using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
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
    public partial class AddProductPage : Page
    {
        private readonly ApiServiceProduct _apiService;
        private readonly Product _product;

        internal AddProductPage(Product product = null)
        {
            InitializeComponent();
            _apiService = new ApiServiceProduct();
            _product = product ?? new Product();

            // Заполнение текстовых полей текущими значениями продукта
            NameTextBox.Text = _product.Name;
            PriceTextBox.Text = _product.Price.ToString();
            VendorCodeTextBox.Text = _product.vendor_code;
            ProductTypeIdTextBox.Text = _product.id_product_type.ToString();
            UnitIdTextBox.Text = _product.id_unit.ToString();
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
                _product.Name = NameTextBox.Text;
                _product.Price = int.Parse(PriceTextBox.Text);
                _product.vendor_code = VendorCodeTextBox.Text;
                _product.id_product_type = int.Parse(ProductTypeIdTextBox.Text);
                _product.id_unit = int.Parse(UnitIdTextBox.Text);

                if (_product.id_Product == 0)
                {
                    await _apiService.AddProductAsync(_product);
                    MessageBox.Show("Продукт успешно добавлен");
                }
                else
                {
                    await _apiService.UpdateProductAsync(_product);
                    MessageBox.Show("Продукт успешно обновлен");
                }

                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении продукта: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}