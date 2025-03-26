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
    /// Логика взаимодействия для AddReceiptAndExpenseDocumentsPage.xaml
    /// </summary>
    public partial class AddReceiptAndExpenseDocumentsPage : Page
    {
        private readonly ApiServiceReceiptAndExpenseDocuments _apiService;
        private readonly Receipt_and_expense_documents _Receipt_and_expense_documents;
        internal AddReceiptAndExpenseDocumentsPage(Receipt_and_expense_documents receipt_And_Expense_Documents = null)
        {
            InitializeComponent();
            _apiService = new ApiServiceReceiptAndExpenseDocuments();
            _Receipt_and_expense_documents = receipt_And_Expense_Documents ?? new Receipt_and_expense_documents();


            DataTextBox.Text = _Receipt_and_expense_documents.date.ToString(); ;
            ReceiptAndexpense_documentsTextBox.Text = _Receipt_and_expense_documents.ReceiptAndexpense_documents.ToString(); ;
            UserIdTextBox.Text = _Receipt_and_expense_documents.id_users.ToString();
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
                _Receipt_and_expense_documents.date = DateTime.Parse(DataTextBox.Text);
                _Receipt_and_expense_documents.ReceiptAndexpense_documents = bool.Parse(ReceiptAndexpense_documentsTextBox.Text);
                _Receipt_and_expense_documents.id_users = int.Parse(UserIdTextBox.Text);

                if (_Receipt_and_expense_documents.id_doc == 0)
                {
                    await _apiService.AddReceipt_and_expense_documentsAsync(_Receipt_and_expense_documents);
                    MessageBox.Show("Документ успешно добавлен");
                }
                else
                {
                    await _apiService.UpdateReceipt_and_expense_documentsAsync(_Receipt_and_expense_documents);
                    MessageBox.Show("Документ успешно обновлен");
                }
                NavigationService.GoBack();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении документа: {ex.Message}");
            }
        }

        private void BackButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.GoBack();
        }
    }
}
