using Microsoft.Win32;
using Spire.Xls;
using System;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;

namespace _99
{
    public partial class Excel : Window
    {
        public Excel()
        {
            InitializeComponent();
        }

        private void Create_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.Worksheets.Clear();
            Worksheet sheet = workbook.Worksheets.Add("новый листик");
            var dataview = grid.ItemsSource as DataView;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|.xls;.xlsx;*.xlsm";
            saveFileDialog.Title = "Куда сохранить Excel-файл?";

            if (saveFileDialog.ShowDialog() == true)
            {
                string filePath = saveFileDialog.FileName;
                workbook.SaveToFile(filePath, FileFormat.Version2010);
                Workbook loadedWorkbook = new Workbook();
                workbook.Worksheets.Clear();
                loadedWorkbook.LoadFromFile(filePath);
                Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
                var loadedDataView = loadedSheet.ExportDataTable().DefaultView;
                grid.ItemsSource = loadedDataView;
            }
        }

        private void ClearGrid()
        {
            try
            {
                DataTable dt = ((DataView)grid.ItemsSource).Table;
                dt.Clear();
                grid.ItemsSource = null;

            }

            catch (Exception ex)
            {
                MessageBox.Show("датагрида и так нет!");
            }
        }

        private void ClearTable()
        {
            try
            {
                DataTable dt = ((DataView)grid.ItemsSource).Table;
                dt.Clear();
                grid.ItemsSource = null;
                grid.ItemsSource = dt.DefaultView;
            }

            catch (Exception ex)
            {
                MessageBox.Show("таблицы и так нет!");
            }
        }

        private void Open_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите Excel-файл для работы с ним";

            if (openFileDialog.ShowDialog() == true)
            {

                string filePath = openFileDialog.FileName;
                Workbook workbook = new Workbook();
                workbook.Worksheets.Clear();
                workbook.LoadFromFile(filePath);
                Worksheet sheet = workbook.Worksheets[0];
                CellRange range = sheet.AllocatedRange;
                var datatable = sheet.ExportDataTable(range, true);
                grid.ItemsSource = datatable.DefaultView;
            }
        }

        private void Add_Column_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                AddColumns();
            }
            catch
            {
                MessageBox.Show("Ошибка при добавлении столбца");
            }
        }

        private void Delete_Row_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DeleteRows();
            }
            catch
            {
                MessageBox.Show("ошибка при удалении строчки");
            }
        }

        private void Clear_DataGrid_Button_Click(object sender, RoutedEventArgs e)
        {
            ClearGrid();
        }

        private void Clear_Table_Button_Click(object sender, RoutedEventArgs e)
        {
            ClearTable();
        }

        private void Send_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите Excel-файл для отправки";

            if (openFileDialog.ShowDialog() == true)
            {
                Send excel = new Send(openFileDialog.FileName);
                excel.ShowDialog();
            }
        }

        private void AddColumns()
        {
            DataTable dt = ((DataView)grid.ItemsSource).Table;
            string name = namecol.Text;

            if (!dt.Columns.Contains(name))
            {
                dt.Columns.Add(name);
            }

            DataView dataView = dt.DefaultView;
            grid.ItemsSource = null;
            grid.ItemsSource = dataView;
        }

        private void DeleteRows()
        {
            if (grid.SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)grid.SelectedItem;
                DataTable dt = ((DataView)grid.ItemsSource).Table;
                dt.Rows.Remove(selectedRow.Row);
            }
            else
            {
                MessageBox.Show("Ошибка при удалении строчки");
            }
        }

        private void Exit_Button_Click(object sender, RoutedEventArgs e)
        {
            var window = GetWindow(this);

            if (window != null)
            {
                window.Close();
            }
        }

        private void Save_Excel_Button_Click(object sender, RoutedEventArgs e)
        {
            if (grid.ItemsSource == null)
            {
                MessageBox.Show("Нет данных для сохранения");
                return;
            }

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            saveFileDialog.Title = "Сохранить Excel-файл";

            try
            {
                if (File.Exists(saveFileDialog.FileName))
                {
                    Workbook workbook = new Workbook();
                    workbook.Worksheets.Clear();
                    Worksheet sheet = workbook.Worksheets[0];
                    var dataview = grid.ItemsSource as DataView;
                    sheet.InsertDataView(dataview, true, 1, 1);
                    workbook.SaveToFile(saveFileDialog.FileName, FileFormat.Version2010);
                }
                else
                {
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        Workbook workbook = new Workbook();
                        workbook.Worksheets.Clear();
                        Worksheet sheet = workbook.Worksheets.Add("новый листик");
                        var dataview = grid.ItemsSource as DataView;
                        sheet.InsertDataView(dataview, true, 1, 1);
                        workbook.SaveToFile(saveFileDialog.FileName, FileFormat.Version2010);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при сохранении: " + ex);
            }
        }
    }
}

    