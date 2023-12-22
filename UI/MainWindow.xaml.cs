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
using System.Dynamic;

namespace DataHandling
{
    public sealed partial class MainWindow : Window
    {
        public PlotModel MyModel { get; private set; }
        public ObservableCollection<dynamic> DataGridCollection { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            MyModel = new PlotModel { Title = "Sample Chart" };
            SetupPlotModel();
            DataGridCollection = new ObservableCollection<dynamic>();
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

        private async void LoadExcelFileAsync(string filePath)
        {
            Stopwatch watch = new Stopwatch();

            watch.Start();
            // Clear existing data
            DataGridCollection.Clear();
            DataDisplayGrid.Columns.Clear();

            try
            {
                // Offloading CPU-bound work to a background thread
                await Task.Run(() =>
                {
                    var dataTable = new DataTable();

                    using var stream = File.OpenRead(filePath);
                    using var workbook = new XLWorkbook(stream);
                    var worksheet = workbook.Worksheets.Worksheet(1); // Assuming the first worksheet

                    // Reading header on the background thread
                    var headerRow = worksheet.FirstRowUsed();
                    foreach (var headerCell in headerRow.CellsUsed())
                    {
                        dataTable.Columns.Add(headerCell.GetValue<string>());
                    }

                    // Reading rows on the background thread
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

                    Debug.WriteLine(watch.ElapsedMilliseconds + " : Done Reading");

                    // Switching back to the UI thread to update the UI
                    DispatcherQueue.TryEnqueue(() =>
                    {
                        foreach (DataColumn column in dataTable.Columns)
                        {
                            DataDisplayGrid.Columns.Add(new DataGridTextColumn
                            {
                                Header = column.ColumnName,
                                Binding = new Binding { Path = new PropertyPath($"[{column.ColumnName}]") }
                            });
                        }

                        DataDisplayGrid.ItemsSource = dataTable.DefaultView;
                    });
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading file: {ex.Message}");
                // Handle exceptions
            }
            watch.Stop();

            Debug.WriteLine(watch.ElapsedMilliseconds + "");
        }



        private void LoadExcelFile(string filePath)
        {
            DataGridCollection.Clear(); // Clear existing data

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.Worksheet(1); // Assuming you want the first worksheet
                var range = worksheet.RangeUsed();

                // Reading Header
                var headerRow = range.FirstRow();
                var headers = new string[headerRow.CellCount()];
                for (int i = 0; i < headerRow.CellCount(); i++)
                {
                    headers[i] = headerRow.Cell(i + 1).Value.ToString();
                }

                // Reading Rows
                foreach (var row in range.RowsUsed().Skip(1)) // Skip header row
                {
                    dynamic expando = new ExpandoObject();
                    var expandoDict = (IDictionary<string, object>)expando;

                    for (int i = 0; i < row.CellCount(); i++)
                    {
                        string header = headers[i];
                        expandoDict[header] = row.Cell(i + 1).Value;
                    }

                    DataGridCollection.Add(expando);
                }
            }

            // Bind to DataGrid
            DataDisplayGrid.ItemsSource = DataGridCollection;
        }

        private async void LoadExcelFileAsync(string filePath)
        {
            DataGridCollection.Clear(); // Clear existing data
            DataDisplayGrid.Columns.Clear();

            try
            {
                // Offloading CPU-bound work to a background thread
                await Task.Run(() =>
                {
                    using var workbook = new XLWorkbook(filePath);
                    var worksheet = workbook.Worksheet(1); // Assuming the first worksheet
                    var totalRows = worksheet.RangeUsed().RowCount();
                    int chunkSize = 1000; // Define your chunk size

                    for (int row = 1; row <= totalRows; row += chunkSize)
                    {
                        int chunkEnd = Math.Min(row + chunkSize - 1, totalRows);
                        DataTable chunkDataTable = ReadExcelChunk(worksheet, row, chunkEnd);

                        // Switching back to the UI thread to update the UI
                        DispatcherQueue.TryEnqueue(() =>
                        {
                            UpdateDataGrid(chunkDataTable);
                        });
                    }
                });
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error reading file: {ex.Message}");
                // Handle exceptions
            }
        }

        private DataTable ReadExcelChunk(IXLWorksheet worksheet, int startRow, int endRow)
        {
            var dataTable = new DataTable();

            // Read headers
            if (startRow == 1)
            {
                var headerRow = worksheet.Row(1);
                foreach (var headerCell in headerRow.CellsUsed())
                {
                    dataTable.Columns.Add(headerCell.GetValue<string>());
                }
                startRow++; // Skip header row for subsequent data
            }

            // Read data rows
            for (int rowNum = startRow; rowNum <= endRow; rowNum++)
            {
                var row = worksheet.Row(rowNum);
                var dataRow = dataTable.NewRow();
                int columnIndex = 0;
                foreach (var cell in row.CellsUsed())
                {
                    dataRow[columnIndex++] = cell.GetValue<string>();
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        private void UpdateDataGrid(DataTable chunkDataTable)
        {
            if (DataDisplayGrid.Columns.Count == 0)
            {
                foreach (DataColumn column in chunkDataTable.Columns)
                {
                    DataDisplayGrid.Columns.Add(new DataGridTextColumn
                    {
                        Header = column.ColumnName,
                        Binding = new Binding { Path = new PropertyPath($"[{column.ColumnName}]") }
                    });
                }
            }

            foreach (DataRow row in chunkDataTable.Rows)
            {
                DataGridCollection.Add(row);
            }

            DataDisplayGrid.ItemsSource = DataGridCollection;
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
                LoadExcelFileAsync(file.Path);
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
