using System;
using System.Collections.Generic;
using System.Windows;

namespace GenerateReportTable
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Список данных
        private List<Report> table;

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
            table = new List<Report>();

            for (int i = 1; i < 16; i++)
            {
                table.Add(new Report(i, marks[i - 1]));
            }

            // Привязка данных
            gridReport.ItemsSource = table;

            btnWord.IsEnabled = true;

            // Размеры
            Height = 370;
            Width = 946;

            gridReport.Visibility = Visibility.Visible;
        }

        private async void btnWord_Click(object sender, RoutedEventArgs e)
        {
            // Создать документ
            WordDoc w = new WordDoc();
            await w.CreateWordAsync();

            // Создать название таблицы
            await w.CreateTableNameAsync();

            // Создать таблицу
            await w.CreateTableAsync(table);

            // Путь для сохранения
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog()
            {
                Filter = "Файлы Word (*.doc; *.docx)|*.doc;*.docx",
                Title = "Выберите место для сохранения документа",
                DefaultExt = "docx",
                OverwritePrompt = false
            };

            bool? result = dlg.ShowDialog();

            if (result ?? true)
            {
                // Сохранение
                await w.SaveAsync(dlg.FileName);

                // Закрыть документ
                await w.CloseAsync();

                MessageBox.Show("Сохранение завершено");
            }
            else
            {
                MessageBox.Show("Сохранение не удалось", "Ошибка");
            }
        }
    }
}
