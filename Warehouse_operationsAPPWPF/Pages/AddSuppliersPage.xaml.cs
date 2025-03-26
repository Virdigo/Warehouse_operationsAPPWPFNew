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
    /// Логика взаимодействия для AddSuppliersPage.xaml
    /// </summary>
    public partial class AddSuppliersPage : Page
    {
        private readonly ApiServiceSuppliers _apiService;
        private readonly Suppliers _supplier;

        public AddSuppliersPage(Suppliers supplier = null)
        {
            InitializeComponent();
            _apiService = new ApiServiceSuppliers();
            _supplier = supplier ?? new Suppliers();

            // Заполнение текстовых полей текущими значениями поставщика
            NameTextBox.Text = _supplier.Name;
            ContactInformationTextBox.Text = _supplier.Contact_Information;
        }
        private async void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка наличия обязательных данных
            if (string.IsNullOrWhiteSpace(NameTextBox.Text) || string.IsNullOrWhiteSpace(ContactInformationTextBox.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Заполнение объекта поставщика новыми значениями
            _supplier.Name = NameTextBox.Text;
            _supplier.Contact_Information = ContactInformationTextBox.Text;

            try
            {
                if (_supplier.id_suppliers == 0)
                {
                    // Если id поставщика равен 0, добавляем нового поставщика
                    await _apiService.AddSuppliersAsync(_supplier);
                    MessageBox.Show("Поставщик успешно добавлен.");
                }
                else
                {
                    // Иначе обновляем существующего поставщика
                    await _apiService.UpdateSuppliersAsync(_supplier);
                    MessageBox.Show("Поставщик успешно обновлен.");
                }

                // Возвращение на страницу списка поставщиков после сохранения
                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении поставщика: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}