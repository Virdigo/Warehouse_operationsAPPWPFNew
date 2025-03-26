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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для ReceiptAndExpenseDocumentsPage.xaml
    /// </summary>
    public partial class ReceiptAndExpenseDocumentsPage : Page
    {
        private readonly ApiServiceReceiptAndExpenseDocuments _apiService;
        private List<Receipt_and_expense_documents> _allReceiptAndExpenseDocuments;
        public ReceiptAndExpenseDocumentsPage()
        {
            InitializeComponent();
            _apiService = new ApiServiceReceiptAndExpenseDocuments();
            LoadReceiptAndExpenseDocuments();
        }
        private async void LoadDocuments_Click(object sender, RoutedEventArgs e)
        {
            await LoadReceiptAndExpenseDocuments();
        }

        private async Task LoadReceiptAndExpenseDocuments()
        {
            try
            {
                _allReceiptAndExpenseDocuments = await _apiService.GetReceiptAndExpenseDocumentsList();
                DocumentsListView.ItemsSource = _allReceiptAndExpenseDocuments;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки документов: {ex.Message}");
            }
        }

        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredProducts = _allReceiptAndExpenseDocuments.AsEnumerable();

           

            // Фильтр по дате
            if (StartDatePicker.SelectedDate.HasValue)
            {
                var startDate = StartDatePicker.SelectedDate.Value;
                filteredProducts = filteredProducts.Where(p => p.date >= startDate);
            }

            if (EndDatePicker.SelectedDate.HasValue)
            {
                var endDate = EndDatePicker.SelectedDate.Value;
                filteredProducts = filteredProducts.Where(p => p.date <= endDate);
            }

            // Обновление ListView
            DocumentsListView.ItemsSource = filteredProducts.ToList();
        }

        private void AddDocumentsButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddReceiptAndExpenseDocumentsPage());
        }

        private void EditDocumentsButton_Click(object sender, RoutedEventArgs e)
        {
            if (DocumentsListView.SelectedItem is Receipt_and_expense_documents selectedReceiptAndExpenseDocuments)
            {
                NavigationService.Navigate(new AddReceiptAndExpenseDocumentsPage(selectedReceiptAndExpenseDocuments));
            }
            else
            {
                MessageBox.Show("Выберите документ для редактирования.");
            }
        }

        private async void DeleteDocumentsButton_Click(object sender, RoutedEventArgs e)
        {
            if (DocumentsListView.SelectedItem is Receipt_and_expense_documents selectedReceiptAndExpenseDocuments)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить документ {selectedReceiptAndExpenseDocuments.id_doc}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiService.DeleteReceipt_and_expense_documentsAsync(selectedReceiptAndExpenseDocuments.id_doc);
                        MessageBox.Show("Документ успешно удален.");
                        await LoadReceiptAndExpenseDocuments();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка при удалении документа: {ex.Message}");
                    }
                }
            }
            else
            {
                MessageBox.Show("Выберите документ для удаления.");
            }
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new SuppliersPage());
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new InformationPage());
        }

        private void excel_Click(object sender, RoutedEventArgs e)
        {
            var ExcelApp = new Excel.Application();

            Excel.Workbook wb = ExcelApp.Workbooks.Add();

            Excel.Worksheet worksheet = ExcelApp.Worksheets.Item[1];

            int indexRows = 1;
            worksheet.Cells[1][indexRows] = "Код документа";
            worksheet.Cells[2][indexRows] = "Дата";
            worksheet.Cells[3][indexRows] = "Приходно или расходный документ";
            worksheet.Cells[4][indexRows] = "Код пользователя";

            var printItems = DocumentsListView.Items;

            foreach (Receipt_and_expense_documents item in printItems)
            {
                worksheet.Cells[1][indexRows + 1] = indexRows;
                worksheet.Cells[2][indexRows + 1] = item.date;
                worksheet.Cells[3][indexRows + 1] = item.ReceiptAndexpense_documents;
                worksheet.Cells[4][indexRows + 1] = item.id_users;

                indexRows++;
            }
            Excel.Range range = worksheet.Range[worksheet.Cells[2][indexRows + 1],
                    worksheet.Cells[7][indexRows + 1]];

            range.ColumnWidth = 20;

            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            ExcelApp.Visible = true;
        }

        private async void PDF_Click(object sender, RoutedEventArgs e)
        {
            _allReceiptAndExpenseDocuments = await _apiService.GetReceiptAndExpenseDocumentsList();
            var ReceiptAndExpenseDocumentsInPDF = _allReceiptAndExpenseDocuments;

            var ReceiptAndExpenseDocumentsApplicationPDF = new Word.Application();

            Word.Document document = ReceiptAndExpenseDocumentsApplicationPDF.Documents.Add();

            Word.Paragraph empParagraph = document.Paragraphs.Add();
            Word.Range empRange = empParagraph.Range;
            empRange.Text = "Документы";
            empRange.Font.Bold = 4;
            empRange.Font.Italic = 4;
            empRange.Font.Color = Word.WdColor.wdColorBlack;
            empRange.InsertParagraphAfter();

            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = tableParagraph.Range;
            Word.Table paymentsTable = document.Tables.Add(tableRange, ReceiptAndExpenseDocumentsInPDF.Count() + 1, 4);
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Word.Range cellRange;

            cellRange = paymentsTable.Cell(1, 1).Range;
            cellRange.Text = "Код документа";
            cellRange = paymentsTable.Cell(1, 2).Range;
            cellRange.Text = "Дата";
            cellRange = paymentsTable.Cell(1, 3).Range;
            cellRange.Text = "Приходно или расходный документ";
            cellRange = paymentsTable.Cell(1, 4).Range;
            cellRange.Text = "Код пользователя";



            paymentsTable.Rows[1].Range.Bold = 1;
            paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 0; i < ReceiptAndExpenseDocumentsInPDF.Count(); i++)
            {
                var ProductCurrent = ReceiptAndExpenseDocumentsInPDF[i];

                cellRange = paymentsTable.Cell(i + 2, 1).Range;
                cellRange.Text = ProductCurrent.id_doc.ToString();

                cellRange = paymentsTable.Cell(i + 2, 2).Range;
                cellRange.Text = ProductCurrent.date.ToString();

                cellRange = paymentsTable.Cell(i + 2, 3).Range;
                cellRange.Text = ProductCurrent.ReceiptAndexpense_documents.ToString();

                cellRange = paymentsTable.Cell(i + 2, 4).Range;
                cellRange.Text = ProductCurrent.id_users.ToString();
            }

            ReceiptAndExpenseDocumentsApplicationPDF.Visible = true;

            document.SaveAs2(@"C:\Users\bpvla\Desktop\Проект в авторизацей\Warehouse_operationsAPPWPF-master\Warehouse_operationsAPPWPF\bin\Debug\ReceiptAndExpenseDocuments.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    }
}