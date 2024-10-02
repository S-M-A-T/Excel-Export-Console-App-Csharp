using System;
using System.Data;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace ExcelExportConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            // Register the encoding provider to support IBM437 encoding
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            // Step 1: Create and populate DataTable
            DataTable dataTable = CreateDataTable();

            // Step 2: Export DataTable to Excel
            ExportToExcel(dataTable, "VehicleRecords.xlsx");

            Console.WriteLine("Data exported to Excel successfully.");
            Console.ReadLine();
        }

        static DataTable CreateDataTable()
        {
            DataTable dataTable = new DataTable("VehicleRecords");

            // Define columns
            dataTable.Columns.Add("Date", typeof(DateTime));
            dataTable.Columns.Add("DriverName", typeof(string));
            dataTable.Columns.Add("VehicleNumber", typeof(string));
            dataTable.Columns.Add("Purpose", typeof(string));
            dataTable.Columns.Add("DepartureTime", typeof(DateTime));
            dataTable.Columns.Add("ArrivalTime", typeof(DateTime));
            dataTable.Columns.Add("TotalMileage", typeof(double));

            // Step 3: Populate DataTable with sample data
            for (int i = 1; i <= 300000; i++)
            {
                DataRow row = dataTable.NewRow();
                row["Date"] = DateTime.Now.AddDays(-i); // Sample dates
                row["DriverName"] = "Driver " + i; // Use string concatenation
                row["VehicleNumber"] = "Vehicle " + i; // Use string concatenation
                row["Purpose"] = "Delivery";
                row["DepartureTime"] = DateTime.Now.AddHours(-i);
                row["ArrivalTime"] = DateTime.Now.AddHours(-i + 2);
                row["TotalMileage"] = i * 10; // Sample mileage

                dataTable.Rows.Add(row);
            }

            return dataTable;
        }

        static void ExportToExcel(DataTable dataTable, string filePath)
        {
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                var worksheet = excelPackage.Workbook.Worksheets.Add("Vehicle Records");

                // Load the DataTable into the Excel worksheet
                worksheet.Cells["A1"].LoadFromDataTable(dataTable, true);

                // Save the Excel file
                FileInfo excelFile = new FileInfo(filePath);
                excelPackage.SaveAs(excelFile);
            }
        }
    }
}
