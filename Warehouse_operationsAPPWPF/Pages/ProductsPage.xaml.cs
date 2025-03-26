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
    public partial class ProductsPage : Page
    {
        private readonly ApiServiceProduct _apiService;
        private List<Product> _allProducts;

        public ProductsPage()
        {
            InitializeComponent();
            _apiService = new ApiServiceProduct();
            LoadProducts();
        }

        private async void LoadProducts_Click(object sender, RoutedEventArgs e)
        {
            await LoadProducts();
        }
        private void FilterTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !Regex.IsMatch(e.Text, @"^[0-9]+$");
        }
        private async Task LoadProducts()
        {
            try
            {
                _allProducts = await _apiService.GetProductsAsync();
                ProductsListView.ItemsSource = _allProducts;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки продуктов: {ex.Message}");
            }
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredProducts = _allProducts.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(NameFilterTextBox.Text))
            {
                var filterText = NameFilterTextBox.Text.ToLowerInvariant();
                filteredProducts = filteredProducts.Where(p => p.Name.ToLowerInvariant().Contains(filterText));
            }

            if (int.TryParse(MinPriceFilterTextBox.Text, out var minPrice))
            {
                filteredProducts = filteredProducts.Where(p => p.Price >= minPrice);
            }

            if (int.TryParse(MaxPriceFilterTextBox.Text, out var maxPrice))
            {
                filteredProducts = filteredProducts.Where(p => p.Price <= maxPrice);
            }

            ProductsListView.ItemsSource = filteredProducts.ToList();
        }

        private void AddProductButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddProductPage());
        }

        private void EditProductButton_Click(object sender, RoutedEventArgs e)
        {
            if (ProductsListView.SelectedItem is Product selectedProduct)
            {
                NavigationService.Navigate(new AddProductPage(selectedProduct));
            }
            else
            {
                MessageBox.Show("Выберите продукт для редактирования.");
            }
        }

        private async void DeleteProductButton_Click(object sender, RoutedEventArgs e)
        {
            if (ProductsListView.SelectedItem is Product selectedProduct)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить продукт {selectedProduct.Name}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiService.DeleteProductAsync(selectedProduct.id_Product);
                        MessageBox.Show("Продукт успешно удален.");
                        await LoadProducts();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении продукта: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите продукт для удаления.");
            }
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new OstatkiPage());
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new Information_about_documents());
        }
    }
}