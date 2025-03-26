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
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для OstatkiPage.xaml
    /// </summary>
    public partial class OstatkiPage : Page
    {
        private readonly ApiServiceOstatki _apiServiceOstatki;
        private List<Ostatki> _allOstatki;
        public OstatkiPage()
        {
            InitializeComponent();
            _apiServiceOstatki = new ApiServiceOstatki();
            LoadOstatki();
        }
        private async void LoadOstatki_Click(object sender, RoutedEventArgs e)
        {
            await LoadOstatki();
        }

        private async Task LoadOstatki()
        {
            try
            {
                _allOstatki = await _apiServiceOstatki.GetOstatkisAsync();
                OstatkiListView.ItemsSource = _allOstatki;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки продуктов: {ex.Message}");
            }
        }
        private void FilterTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Разрешаем только ввод цифр и знак минус
            e.Handled = !Regex.IsMatch(e.Text, @"^[0-9]+$");
        }
        private void FilterButton_Click(object sender, RoutedEventArgs e)
        {
            var filteredOstatki = _allOstatki.AsEnumerable();

          

            if (int.TryParse(MinQuantityFilterTextBox.Text, out var minOstatki))
            {
                filteredOstatki = filteredOstatki.Where(p => p.Quantity_Ostatki >= minOstatki);
            }

            if (int.TryParse(MaxQuantityFilterTextBox.Text, out var maxOstatki))
            {
                filteredOstatki = filteredOstatki.Where(p => p.Quantity_Ostatki <= maxOstatki);
            }

            OstatkiListView.ItemsSource = filteredOstatki.ToList();
        }

        private void AddOstatkiButton_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new AddOstatkiPage());
        }

        private void EditOstatkiButton_Click(object sender, RoutedEventArgs e)
        {
            if (OstatkiListView.SelectedItem is Ostatki selectedOstatki)
            {
                NavigationService.Navigate(new AddOstatkiPage(selectedOstatki));
            }
            else
            {
                MessageBox.Show("Выберите продукт для редактирования.");
            }
        }

        private async void DeleteOstatkiButton_Click(object sender, RoutedEventArgs e)
        {
            if (OstatkiListView.SelectedItem is Ostatki selectedOstatki)
            {
                var result = MessageBox.Show($"Вы уверены, что хотите удалить остатки {selectedOstatki.id_warehouses}?", "Подтверждение удаления", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    try
                    {
                        await _apiServiceOstatki.DeleteOstatkiAsync(selectedOstatki.id_Ostatki);
                        MessageBox.Show("Продукт успешно удален.");
                        await LoadOstatki();
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

        private async void PDF_Click(object sender, RoutedEventArgs e)
        {
            _allOstatki = await _apiServiceOstatki.GetOstatkisAsync();
            var OstatkiInPDF = _allOstatki;

            var OstatkiApplicationPDF = new Word.Application();

            Word.Document document = OstatkiApplicationPDF.Documents.Add();

            Word.Paragraph empParagraph = document.Paragraphs.Add();
            Word.Range empRange = empParagraph.Range;
            empRange.Text = "Ostatki";
            empRange.Font.Bold = 4;
            empRange.Font.Italic = 4;
            empRange.Font.Color = Word.WdColor.wdColorBlack;
            empRange.InsertParagraphAfter();

            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = tableParagraph.Range;
            Word.Table paymentsTable = document.Tables.Add(tableRange, OstatkiInPDF.Count() + 1, 4);
            paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

            Word.Range cellRange;

            cellRange = paymentsTable.Cell(1, 1).Range;
            cellRange.Text = "Код остатков";
            cellRange = paymentsTable.Cell(1, 2).Range;
            cellRange.Text = "Код склада";
            cellRange = paymentsTable.Cell(1, 3).Range;
            cellRange.Text = "Код товара";
            cellRange = paymentsTable.Cell(1, 4).Range;
            cellRange.Text = "Количество остатков";



            paymentsTable.Rows[1].Range.Bold = 1;
            paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            for (int i = 0; i < OstatkiInPDF.Count(); i++)
            {
                var ProductCurrent = OstatkiInPDF[i];

                cellRange = paymentsTable.Cell(i + 2, 1).Range;
                cellRange.Text = ProductCurrent.id_Ostatki.ToString();

                cellRange = paymentsTable.Cell(i + 2, 2).Range;
                cellRange.Text = ProductCurrent.id_warehouses.ToString();

                cellRange = paymentsTable.Cell(i + 2, 3).Range;
                cellRange.Text = ProductCurrent.id_Product.ToString();

                cellRange = paymentsTable.Cell(i + 2, 4).Range;
                cellRange.Text = ProductCurrent.Quantity_Ostatki.ToString();
            }

            OstatkiApplicationPDF.Visible = true;

            document.SaveAs2(@"C:\Users\bpvla\Desktop\Проект в авторизацей\Warehouse_operationsAPPWPF-master\Warehouse_operationsAPPWPF\bin\Debug\Ostatki.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }

        private void excel_Click(object sender, RoutedEventArgs e)
        {
            var ExcelApp = new Excel.Application();

            Excel.Workbook wb = ExcelApp.Workbooks.Add();

            Excel.Worksheet worksheet = ExcelApp.Worksheets.Item[1];

            int indexRows = 1;
            worksheet.Cells[1][indexRows] = "Код остатков";
            worksheet.Cells[2][indexRows] = "Код склада";
            worksheet.Cells[3][indexRows] = "Код товара";
            worksheet.Cells[4][indexRows] = "Количество остатков";

            var printItems = OstatkiListView.Items;

            foreach (Ostatki item in printItems)
            {
                worksheet.Cells[1][indexRows + 1] = indexRows;
                worksheet.Cells[2][indexRows + 1] = item.id_warehouses;
                worksheet.Cells[3][indexRows + 1] = item.id_Product;
                worksheet.Cells[4][indexRows + 1] = item.Quantity_Ostatki;

                indexRows++;
            }
            Excel.Range range = worksheet.Range[worksheet.Cells[2][indexRows + 1],
                    worksheet.Cells[7][indexRows + 1]];

            range.ColumnWidth = 20;

            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            ExcelApp.Visible = true;
        }

        private void BtnArrowLeft_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ProductsPage());
        }

        private void BtnArrowRight_Click(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new WarehousePage());
        }
    }
}
    

