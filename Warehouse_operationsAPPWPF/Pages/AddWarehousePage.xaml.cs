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
    /// Логика взаимодействия для AddWarehousePage.xaml
    /// </summary>
    public partial class AddWarehousePage : Page
    {
        private readonly ApiServiceWarehouses _apiService;
        private readonly Warehouses _warehouse;

        internal AddWarehousePage(Warehouses warehouse = null)
        {
            InitializeComponent();
            _apiService = new ApiServiceWarehouses();
            _warehouse = warehouse ?? new Warehouses();

            // Заполнение текстовых полей текущими значениями склада
            NameTextBox.Text = _warehouse.Name;
            AddressTextBox.Text = _warehouse.address;
            UserIdTextBox.Text = _warehouse.id_users.ToString();
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
                _warehouse.Name = NameTextBox.Text;
                _warehouse.address = AddressTextBox.Text;
                _warehouse.id_users = int.Parse(UserIdTextBox.Text);

                if (_warehouse.id_warehouses == 0)
                {
                    await _apiService.AddWarehousesAsync(_warehouse);
                    MessageBox.Show("Склад успешно добавлен");
                }
                else
                {
                    await _apiService.UpdateWarehousesAsync(_warehouse);
                    MessageBox.Show("Склад успешно обновлен");
                }
                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении склада: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}

