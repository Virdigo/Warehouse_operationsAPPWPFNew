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
using Warehouse_operationsAPPWPF.Pages;
using Warehouse_operationsAPPWPF.Pages.Accountant;
using Warehouse_operationsAPPWPF.Pages.WorkWarehouse;
using Warehouse_operationsAPPWPF.Services;

namespace Warehouse_operationsAPPWPF
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly ApiServiceLogin _apiService;

        public MainWindow()
        {
            InitializeComponent();
            MainFrame.Navigate(new OstatkiPage());
        }

        // Свернуть окно
        private void BtnMinimize_Click_1(object sender, RoutedEventArgs e)
        {
            Application.Current.MainWindow.WindowState = WindowState.Minimized;
        }

        // Развернуть или вернуть окно в нормальное состояние
        private void BtnMaximize_Click_1(object sender, RoutedEventArgs e)
        {
            if (Application.Current.MainWindow.WindowState == WindowState.Normal)
                Application.Current.MainWindow.WindowState = WindowState.Maximized;
            else
                Application.Current.MainWindow.WindowState = WindowState.Normal;
        }

        // Закрыть окно
        private void BtnClose_Click_1(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        // Переход на главную страницу меню
        private void BtnMenu_Click(object sender, RoutedEventArgs e)
        {
            if (MainFrame.Content is LoginPage)
            {
                MessageBox.Show("Сначала авторизируйтесь");
            }

            if (MainFrame.Content is OstatkiPageAccountant)
            {
                return;
            }
            if (MainFrame.Content is ProductsPageAccountant)
            {
                return;
            }
            if (MainFrame.Content is OstatkiPageWarehouse)
            {
                return;
            }
            if (MainFrame.Content is ProductsPageWarehouse)
            {
                return;
            }
            if (MainFrame.Content is WarehousePageWarehouse)
            {
                return;
            }
            else
            {
                MainFrame.Navigate(new PageMenu());
            }
        }

        // Кнопка выхода
        private void logout_Click(object sender, RoutedEventArgs e)
        {
            // Если мы уже на странице логина, ничего не делаем
            if (MainFrame.Content is LoginPage)
            {
                return;
            }


            // Запрашиваем подтверждение выхода
            var result = MessageBox.Show("Вы уверены, что хотите выйти?", "Предупреждение", MessageBoxButton.YesNo, MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                // После выхода, переходим на страницу логина
                MainFrame.Navigate(new LoginPage());
            }
        }
    }
}