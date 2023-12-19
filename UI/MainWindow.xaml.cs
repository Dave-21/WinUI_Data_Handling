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

        public MainWindow()
        {
            this.InitializeComponent();

            SetupPlotModel();
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

            // Add the series to the PlotModel
            MyModel.Series.Add(lineSeries);

            // Assign the model to the view
            MyPlotView.Model = MyModel;
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
            try
            {
                var picker = new FileOpenPicker();
                picker.FileTypeFilter.Add(".xlsx");
                picker.FileTypeFilter.Add(".csv");
                picker.SuggestedStartLocation = PickerLocationId.DocumentsLibrary;

                WinRT.Interop.InitializeWithWindow.Initialize(picker, App.MainWindowHandle);

                StorageFile file = await picker.PickSingleFileAsync();
                if (file != null)
                {
                    ReadAndDisplayData(file);
                }
            }
            catch (Exception ex)
            {
                // Handle or log the exception
            }
        }

        /* Data Handling */
        private async void ReadAndDisplayData(StorageFile file)
        {
            string fileExtension = file.FileType.ToLower();
            DataTable dataTable = new DataTable();

            switch (fileExtension)
            {
                case ".xlsx":
                    dataTable = await ReadExcelFile(file);
                    break;
                case ".csv":
                    dataTable = await ReadCsvFile(file);
                    break;
            }

            DataDisplayGrid.ItemsSource = dataTable.DefaultView;
        }

        private async Task<DataTable> ReadExcelFile(StorageFile file)
        {
            DataTable dt = new DataTable();
            using (var stream = await file.OpenStreamForReadAsync())
            {
                var workbook = new XLWorkbook(stream);
                var worksheet = workbook.Worksheet(1); // Assuming the first worksheet
                                                       // Convert worksheet to DataTable (see implementation below)
                dt = ConvertWorksheetToDataTable(worksheet);
            }
            return dt;
        }

        private async Task<DataTable> ReadCsvFile(StorageFile file)
        {
            DataTable dataTable = new DataTable();
            using (var stream = await file.OpenStreamForReadAsync())
            using (var reader = new StreamReader(stream))
            {
                bool isFirstRow = true;
                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    var values = line.Split(',');

                    if (isFirstRow)
                    {
                        // Assuming the first row contains column headers
                        foreach (var header in values)
                        {
                            dataTable.Columns.Add(header);
                        }
                        isFirstRow = false;
                    }
                    else
                    {
                        dataTable.Rows.Add(values);
                    }
                }
            }
            return dataTable;
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
