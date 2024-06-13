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

namespace _99
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Create_Word_Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Open_Word_Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Create_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            Excel excel = new Excel();
            excel.Show();
            this.Close();
        }

        private void Open_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            
        }
    }
}
