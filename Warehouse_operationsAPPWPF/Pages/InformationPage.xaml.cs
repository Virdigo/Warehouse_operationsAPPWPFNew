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
    /// Логика взаимодействия для InformationPage.xaml
    /// </summary>
    public partial class InformationPage : Page
    {
        private readonly ApiServiceInformation _apiService;
        private List<Information_about_documents> _allInformation;

        public InformationPage()
        {
            InitializeComponent();
            _apiService = new ApiServiceInformation();
            LoadInformation();
        }

        private async void LoadInformation_Click(object sender, RoutedEventArgs e)
        {
            await LoadInformation();
        }

        private async Task LoadInformation()
        {
            try
            {
                _allInformation = await _apiService.GetInformationAsync();
                InformationListView.ItemsSource = _allInformation;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки информации: {ex.Message}");
            }
        }
        private void FilterTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только ввод цифр и знак минус
            e.Handled = !Regex.IsMatch(e.Text, @"^[0-9]+$");
        }
        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredInformation = _allInformation.AsEnumerable();

            if (int.TryParse(MinPriceFilterTextBox.Text, out var minPrice))
            {
                filteredInformation = filteredInformation.Where(p => p.Price >= minPrice);
            }

            if (int.TryParse(MaxPriceFilterTextBox.Text, out var maxPrice))
            {
                filteredInformation = filteredInformation.Where(p => p.Price <= maxPrice);
            }

            InformationListView.ItemsSource = filteredInformation.ToList();
        }

        private void AddInformationButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddInformationPage());
        }

        private void EditInformationButton_Click(object sender, RoutedEventArgs e)
        {
            if (InformationListView.SelectedItem is Information_about_documents selectedInformation)
            {
                NavigationService.Navigate(new AddInformationPage(selectedInformation));
            }
            else
            {
                MessageBox.Show("Выберите информацию для редактирования.");
            }
        }

        private async void DeleteInformationButton_Click(object sender, RoutedEventArgs e)
        {
            if (InformationListView.SelectedItem is Information_about_documents selectedInformation)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить информацию с ID {selectedInformation.id_inf_doc}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiService.DeleteInformationAsync(selectedInformation.id_inf_doc);
                        MessageBox.Show("Информация успешно удалена.");
                        await LoadInformation();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении информации: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите информацию для удаления.");
            }
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ReceiptAndExpenseDocumentsPage());
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ProductsPage());
        }
    }
}