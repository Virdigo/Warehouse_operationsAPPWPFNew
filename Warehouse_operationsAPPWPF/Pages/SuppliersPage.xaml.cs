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
using Warehouse_operationsAPPWPF.Models;
using Warehouse_operationsAPPWPF.Services;

namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для SuppliersPage.xaml
    /// </summary>
    public partial class SuppliersPage : Page
    {
        private readonly ApiServiceSuppliers _apiService;
        private List<Suppliers> _allSuppliers;

        public SuppliersPage()
        {
            InitializeComponent();
            _apiService = new ApiServiceSuppliers();
            LoadSuppliers();
        }

        private async void LoadSuppliers_Click(object sender, RoutedEventArgs e)
        {
            await LoadSuppliers();
        }

        private async Task LoadSuppliers()
        {
            try
            {
                _allSuppliers = await _apiService.GetSuppliersAsync();
                SuppliersListView.ItemsSource = _allSuppliers;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки поставщиков: {ex.Message}");
            }
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredSuppliers = _allSuppliers.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(NameFilterTextBox.Text))
            {
                var filterText = NameFilterTextBox.Text.ToLowerInvariant();
                filteredSuppliers = filteredSuppliers.Where(s => s.Name.ToLowerInvariant().Contains(filterText));
            }

            SuppliersListView.ItemsSource = filteredSuppliers.ToList();
        }

        private void AddSupplierButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddSuppliersPage());
        }

        private void EditSupplierButton_Click(object sender, RoutedEventArgs e)
        {
            if (SuppliersListView.SelectedItem is Suppliers selectedSupplier)
            {
                NavigationService.Navigate(new AddSuppliersPage(selectedSupplier));
            }
            else
            {
                MessageBox.Show("Выберите поставщика для редактирования.");
            }
        }

        private async void DeleteSupplierButton_Click(object sender, RoutedEventArgs e)
        {
            if (SuppliersListView.SelectedItem is Suppliers selectedSupplier)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить поставщика {selectedSupplier.Name}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiService.DeleteSuppliersAsync(selectedSupplier.id_suppliers);
                        MessageBox.Show("Поставщик успешно удален.");
                        await LoadSuppliers();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении поставщика: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите поставщика для удаления.");
            }
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new WarehousePage());
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ReceiptAndExpenseDocumentsPage());
        }
    }
}
