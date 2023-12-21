using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Controls;
using Microsoft.UI.Xaml.Controls.Primitives;
using Microsoft.UI.Xaml.Data;
using Microsoft.UI.Xaml.Input;
using Microsoft.UI.Xaml.Media;
using Microsoft.UI.Xaml.Navigation;
using System;
using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using ClosedXML.Excel;
using H.OxyPlot.WinUI;
using OxyPlot;
using OxyPlot.Series;
using Windows.Storage.Pickers;
using Windows.Storage;
using System.Threading.Tasks;
using System.Data;
using System.Diagnostics;
using System.Timers;
using CommunityToolkit.WinUI.UI.Controls;
using Microsoft.UI.Dispatching;
using WinRT.Interop;

// To learn more about WinUI, the WinUI project structure,
// and more about our project templates, see: http://aka.ms/winui-project-info.

namespace DataHandling
{
    /// <summary>
    /// An empty window that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainWindow : Window
    {
        public PlotModel MyModel { get; private set; }
        public DataTable MyDataTable { get; set; }
        private DispatcherQueue dispatcherQueue;

        public MainWindow()
        {
            this.InitializeComponent();

            dispatcherQueue = this.DispatcherQueue;

            SetupPlotModel();
            //LoadSampleData();
        }

        private void SetupDataGridColumns(DataTable dataTable)
        {
            DataDisplayGrid.Columns.Clear();

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var dataGridColumn = new CommunityToolkit.WinUI.UI.Controls.DataGridTextColumn
                {
                    Header = dataTable.Columns[i].ColumnName,
                    Binding = new Binding
                    {
                        Path = new PropertyPath($"[{i}]")
                    }
                };

                DataDisplayGrid.Columns.Add(dataGridColumn);
            }
        }


        private void PrintDataTable(DataTable dataTable)
        {
            if (dataTable == null)
            {
                Console.WriteLine("DataTable is null");
                return;
            }

            // Print column headers
            string columnHeaders = string.Join("\t", dataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName));
            Debug.WriteLine(columnHeaders);

            // Print each row of data
            foreach (DataRow row in dataTable.Rows)
            {
                string rowString = string.Join("\t", row.ItemArray.Select(item => item.ToString()));
                Debug.WriteLine(rowString);
            }
        }

        private void LoadSampleData()
        {
            DataTable sampleTable = new DataTable();
            sampleTable.Columns.Add("Column1", typeof(string));
            sampleTable.Columns.Add("Column2", typeof(string));

            sampleTable.Rows.Add("Row1Col1", "Row1Col2");
            sampleTable.Rows.Add("Row2Col1", "Row2Col2");

            PrintDataTable(sampleTable);
            DataDisplayGrid.ItemsSource = null;

            SetupDataGridColumns(sampleTable);

            // Set the ItemsSource of the DataGrid to the DefaultView of the DataTable
            DataDisplayGrid.ItemsSource = sampleTable.DefaultView;
        }

        private void SetupPlotModel()
        {
            MyModel = new PlotModel { Title = "Sample Chart" };

            // Create a line series (you can create other types of series as needed)
            var lineSeries = new LineSeries
            {
                MarkerType = MarkerType.Circle,
                MarkerSize = 4,
                MarkerStroke = OxyColors.White
            };

            // Add sample data to the line series
            lineSeries.Points.Add(new DataPoint(0, 0));
            lineSeries.Points.Add(new DataPoint(10, 10));
            lineSeries.Points.Add(new DataPoint(20, 15));
            lineSeries.Points.Add(new DataPoint(30, 25));

            // Add the series to the PlotModels
            MyModel.Series.Add(lineSeries);

            // Assign the model to the view
            //MyPlotView.Model = MyModel;
        }

/*        private void ReadAndDisplayExcelData(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // or a specific sheet name
                var data = ConvertWorksheetToDataTable(worksheet);
                DataDisplayGrid.ItemsSource = data.DefaultView;
            }
        }*/

        private void NewFile_Click(object sender, RoutedEventArgs e)
        {
            // Logic for creating a new file
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

            // Initialize Picker with the window handle
            IntPtr hwnd = WindowNative.GetWindowHandle(this);
            InitializeWithWindow.Initialize(picker, hwnd);

            StorageFile file = await picker.PickSingleFileAsync();
            if (file != null)
            {
                ReadAndDisplayData(file);
            }
        }

        private async void ReadAndDisplayData(StorageFile file)
        {
            try
            {
                var dataTable = await LoadDataAsync(file);
                SetupDataGrid(dataTable);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading file: {ex.Message}");
            }
        }

        private async Task<DataTable> LoadDataAsync(StorageFile file)
        {
            string fileExtension = file.FileType.ToLower();
            switch (fileExtension)
            {
                case ".xlsx":
                    return await ReadExcelFile(file);
                case ".csv":
                    return await ReadCsvFile(file);
                default:
                    throw new InvalidOperationException("Unsupported file format");
            }
        }

        private async Task<DataTable> ReadExcelFile(StorageFile file)
        {
            var dataTable = new DataTable();

            using (var stream = await file.OpenStreamForReadAsync())
            {
                var workbook = new XLWorkbook(stream);
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
            }

            return dataTable;
        }

        private async Task<DataTable> ReadCsvFile(StorageFile file)
        {
            var dataTable = new DataTable();
            using (var stream = await file.OpenStreamForReadAsync())
            using (var reader = new StreamReader(stream))
            {
                string[] headers = reader.ReadLine().Split(',');
                foreach (var header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                while (!reader.EndOfStream)
                {
                    var values = reader.ReadLine().Split(',');
                    dataTable.Rows.Add(values);
                }
            }

            return dataTable;
        }

        private void SetupDataGrid(DataTable dataTable)
        {
            DataDisplayGrid.Columns.Clear();

            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                var column = dataTable.Columns[i];
                DataDisplayGrid.Columns.Add(new DataGridTextColumn
                {
                    Header = column.ColumnName,
                    // Use the column index for binding
                    Binding = new Binding { Path = new PropertyPath($"[{i}]") }
                });
            }

            DataDisplayGrid.ItemsSource = dataTable.DefaultView;
        }

        private DataTable ConvertWorksheetToDataTable(IXLWorksheet worksheet)
        {
            DataTable dataTable = new DataTable();
            bool firstRow = true;

            foreach (IXLRow row in worksheet.RowsUsed())
            {
                // Use the first row to add columns to DataTable
                if (firstRow)
                {
                    foreach (IXLCell cell in row.Cells())
                    {
                        dataTable.Columns.Add(cell.Value.ToString());
                    }
                    firstRow = false;
                }
                else
                {
                    // Add rows to DataTable
                    dataTable.Rows.Add(row.Cells().Select(c => c.Value.ToString()).ToArray());
                }
            }

            return dataTable;
        }
    }
}
