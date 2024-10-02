# Excel Export Console App

## Project Description
The Excel Export Console App is a C# console application that allows users to generate Excel files from large datasets efficiently. It uses the EPPlus library to create an Excel file from a DataTable containing vehicle record information. The application is designed to handle large amounts of data, making it suitable for exporting heavy datasets with potentially hundreds of thousands of rows.

## Features
- **Data Generation**: Creates a DataTable with sample vehicle records, including fields like Date, Driver Name, Vehicle Number, Purpose, Departure Time, Arrival Time, and Total Mileage.
- **Excel Export**: Utilizes the EPPlus library to export the DataTable to an Excel file, making it easy to work with and analyze data in Excel.
- **Scalability**: Capable of handling large datasets (e.g., 300,000 rows) efficiently.

## Installation Instructions
### Prerequisites
Ensure that you have .NET Framework installed on your system.

### Clone the Repository
```bash
git clone https://github.com/yourusername/ExcelExportConsoleApp.git
Install EPPlus Library
Use NuGet Package Manager to install the EPPlus library. You can do this by running the following command in the Package Manager Console:           Build the Solution
Build the project to restore any missing dependencies.
