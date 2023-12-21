using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Data;
using System;
using System.Data;
using System.IO;
using System.Threading.Tasks;
using ClosedXML.Excel;
using OxyPlot;
using OxyPlot.Series;
using Windows.Storage;
using Windows.Storage.Pickers;
using CommunityToolkit.WinUI.UI.Controls;
using WinRT.Interop;
using System.Diagnostics;
using System.Linq;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace DataHandling
{
    public sealed partial class MainWindow : Window
    {
        public PlotModel MyModel { get; private set; }

        public MainWindow()
        {
            InitializeComponent();
            MyModel = new PlotModel { Title = "Sample Chart" };
            SetupPlotModel();
        }

        private void SetupPlotModel()
        {
            // Initialize and configure the plot model here
            // Example: Adding a line series to the plot model
            var lineSeries = new LineSeries
            {
                MarkerType = MarkerType.Circle,
                MarkerSize = 4,
                MarkerStroke = OxyColors.White
            };
            lineSeries.Points.Add(new DataPoint(0, 0));
            lineSeries.Points.Add(new DataPoint(10, 10));
            lineSeries.Points.Add(new DataPoint(20, 15));
            lineSeries.Points.Add(new DataPoint(30, 25));
            MyModel.Series.Add(lineSeries);
        }

        private async void OpenFile_Click(object sender, RoutedEventArgs e)
        {
            var picker = new FileOpenPicker
            {
                ViewMode = PickerViewMode.List,
                SuggestedStartLocation = PickerLocationId.DocumentsLibrary
            };
            picker.FileTypeFilter.Add(".xlsx");
            picker.FileTypeFilter.Add(".csv");

            IntPtr hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                await ReadAndDisplayData(file);
            }
        }

        public ObservableCollection<Dictionary<string, object>> ConvertDataTable(DataTable dt)
        {
            var rows = new ObservableCollection<Dictionary<string, object>>();
            foreach (DataRow row in dt.Rows)
            {
                var dict = new Dictionary<string, object>();
                foreach (DataColumn col in dt.Columns)
                {
                    dict[col.ColumnName] = row[col];
                }
                rows.Add(dict);
            }
            return rows;
        }

        private async Task ReadAndDisplayData(StorageFile file)
        {
            try
            {
                var dataTable = await LoadDataAsync(file);
                SetupDataGrid(dataTable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading file: {ex.Message}");
                //((CommunityToolkit.WinUI.UI.Controls.DataGrid)DataDisplayGrid).
            }
        }

        private async Task<DataTable> LoadDataAsync(StorageFile file)
        {
            string fileExtension = file.FileType.ToLower();
            return fileExtension switch
            {
                ".xlsx" => await ReadExcelFile(file),
                ".csv" => await ReadCsvFile(file),
                _ => throw new InvalidOperationException("Unsupported file format"),
            };
        }

        private async Task<DataTable> ReadExcelFile(StorageFile file)
        {
            var dataTable = new DataTable();
            using var stream = await file.OpenStreamForReadAsync();
            using var workbook = new XLWorkbook(stream);
            var worksheet = workbook.Worksheet(1);

            var headerRow = worksheet.FirstRowUsed();
            foreach (var headerCell in headerRow.CellsUsed())
            {
                dataTable.Columns.Add(headerCell.GetValue<string>());
            }

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var dataRow = dataTable.NewRow();
                int columnIndex = 0;
                foreach (var cell in row.Cells())
                {
                    dataRow[columnIndex++] = cell.GetValue<string>();
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private async Task<DataTable> ReadCsvFile(StorageFile file)
        {
            var dataTable = new DataTable();
            using var stream = await file.OpenStreamForReadAsync();
            using var reader = new StreamReader(stream);

            string[] headers = reader.ReadLine().Split(',');
            foreach (var header in headers)
            {
                dataTable.Columns.Add(header.Trim());
            }

            while (!reader.EndOfStream)
            {
                var values = reader.ReadLine().Split(',');
                dataTable.Rows.Add(values);
            }

            return dataTable;
        }

        private void SetupDataGrid(DataTable dataTable)
        {
            DataDisplayGrid.Columns.Clear();
            foreach (DataColumn column in dataTable.Columns)
            {
                DataDisplayGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.ColumnName,
                    Binding = new Binding { Path = new PropertyPath($"[{column.ColumnName}]") }
                });
            }
            DataDisplayGrid.ItemsSource = ConvertDataTable(dataTable);
        }
    }
}
