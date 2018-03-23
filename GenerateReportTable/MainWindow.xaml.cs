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

namespace GenerateReportTable
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnCreate_Click(object sender, RoutedEventArgs e)
        {
            // Для оценок
            List<int> marks = new List<int>();

            Random r = new Random();
            
            for (int i = 0; i < 15; i++)
            {
                marks.Add(r.Next(4, 6));
            }

            // Количество троек
            int count = r.Next(2, 3);

            for (int i = 0; i < count; i++)
            {
                marks[r.Next(0, 15)] = 3;
            }

            // Создание и заполнение списка для создания таблички
            List<Report> table = new List<Report>();

            for (int i = 1; i < 16; i++)
            {
                table.Add(new Report(i, marks[i - 1]));
            }
        }
    }
}
