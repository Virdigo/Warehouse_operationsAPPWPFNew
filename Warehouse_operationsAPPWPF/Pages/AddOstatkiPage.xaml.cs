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
    /// Логика взаимодействия для AddOstatkiPage.xaml
    /// </summary>
    public partial class AddOstatkiPage : Page
    {
        private readonly ApiServiceOstatki _apiServiceOstatki;
        private readonly Ostatki _ostatki;
        internal AddOstatkiPage(Ostatki ostatki = null)
        {
            InitializeComponent();
            _apiServiceOstatki = new ApiServiceOstatki();
            _ostatki = ostatki ?? new Ostatki();

            // Заполнение текстовых полей текущими значениями продукта
            WarehousesIdTextBox.Text = _ostatki.id_warehouses.ToString();
            ProductIdTextBox.Text = _ostatki.id_Product.ToString();
            Quantity_OstatkiTextBox.Text = _ostatki.Quantity_Ostatki.ToString();

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
                _ostatki.id_warehouses = int.Parse(WarehousesIdTextBox.Text);
                _ostatki.id_Product = int.Parse(ProductIdTextBox.Text);
                _ostatki.Quantity_Ostatki = int.Parse(Quantity_OstatkiTextBox.Text);

                if (_ostatki.id_Ostatki == 0)
                {
                    await _apiServiceOstatki.AddOstatkiAsync(_ostatki);
                    MessageBox.Show("Остатки успешно добавлены");
                }
                else
                {
                    await _apiServiceOstatki.UpdateOstatkiAsync(_ostatki);
                    MessageBox.Show("Остатки успешно обновлены");
                }

                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении остатков: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}
