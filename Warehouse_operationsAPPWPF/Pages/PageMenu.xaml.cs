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

namespace Warehouse_operationsAPPWPF.Pages
{
    /// <summary>
    /// Логика взаимодействия для PageMenu.xaml
    /// </summary>
    public partial class PageMenu : Page
    {
        public PageMenu()
        {
            InitializeComponent();
        }

        private void Button_ClickProduct(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ProductsPage());
        }

        private void Button_ClickOstatki(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new OstatkiPage()); 
        }

        private void Button_ClickWarehouse(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new WarehousePage());
        }

        private void Button_ClickSuppliers(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new SuppliersPage());
        }

        private void Button_ClickInformation(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new InformationPage()); 
        }

        private void Button_ClickDocuments(object sender, RoutedEventArgs e)
        {
            NavigationService.Navigate(new ReceiptAndExpenseDocumentsPage());
        }
    }
}
