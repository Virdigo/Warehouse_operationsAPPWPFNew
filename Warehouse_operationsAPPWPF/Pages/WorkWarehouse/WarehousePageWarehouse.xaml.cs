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

namespace Warehouse_operationsAPPWPF.Pages.WorkWarehouse
{
    /// <summary>
    /// Логика взаимодействия для WarehousePageWarehouse.xaml
    /// </summary>
    public partial class WarehousePageWarehouse : Page
    {
        private readonly ApiServiceWarehouses _apiService;
        private List<Warehouses> _allWarehouses;

        public WarehousePageWarehouse()
        {
            InitializeComponent();
            _apiService = new ApiServiceWarehouses();
            LoadWarehouses();
        }

        private async void LoadWarehouses_Click(object sender, RoutedEventArgs e)
        {
            await LoadWarehouses();
        }

        private async Task LoadWarehouses()
        {
            try
            {
                _allWarehouses = await _apiService.GetWarehousesAsync();
                WarehousesListView.ItemsSource = _allWarehouses;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки складов: {ex.Message}");
            }
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredWarehouses = _allWarehouses.AsEnumerable();

            if (!string.IsNullOrWhiteSpace(NameFilterTextBox.Text))
            {
                var filterText = NameFilterTextBox.Text.ToLowerInvariant();
                filteredWarehouses = filteredWarehouses.Where(p => p.Name.ToLowerInvariant().Contains(filterText));
            }

            if (!string.IsNullOrWhiteSpace(AddressFilterTextBox.Text))
            {
                var filterText = AddressFilterTextBox.Text.ToLowerInvariant();
                filteredWarehouses = filteredWarehouses.Where(p => p.address.ToLowerInvariant().Contains(filterText));
            }

            WarehousesListView.ItemsSource = filteredWarehouses.ToList();
        }

        private void AddWarehouseButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddWarehousePage());
        }

        private void EditWarehouseButton_Click(object sender, RoutedEventArgs e)
        {
            if (WarehousesListView.SelectedItem is Warehouses selectedWarehouse)
            {
                NavigationService.Navigate(new AddWarehousePage(selectedWarehouse));
            }
            else
            {
                MessageBox.Show("Выберите склад для редактирования.");
            }
        }

        private async void DeleteWarehouseButton_Click(object sender, RoutedEventArgs e)
        {
            if (WarehousesListView.SelectedItem is Warehouses selectedWarehouse)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить склад {selectedWarehouse.Name}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiService.DeleteWarehousesAsync(selectedWarehouse.id_warehouses);
                        MessageBox.Show("Склад успешно удален.");
                        await LoadWarehouses();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении склада: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите склад для удаления.");
            }
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ProductsPageWarehouse());
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new OstatkiPageWarehouse());
        }
    }
}