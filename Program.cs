using System;
using System.Data;
using Spire.Xls;

public class Program
{
    public static void Main()
    {
        // Path to the XLSB file
        string filePath = @"C:\Users\Honey.Kumar\OneDrive - Shell\Documents\Honey Kumar\workspace\xlsbfile.xlsb";

        // Load the XLSB file using Spire.XLS
        Workbook workbook = new Workbook();
        workbook.LoadFromFile(filePath);

        // Get the "Blank 3-U" worksheet
        Worksheet worksheet = workbook.Worksheets["Blank 3-U"];
        DataTable dt = worksheet.ExportDataTable(worksheet.AllocatedRange, false, true);
        int count= 0;
        foreach (DataRow row in dt.Rows)
        {
            count++;
            if (count == 6)
            {
            // Loop through each column in the DataRow
                foreach (DataColumn column in dt.Columns)
                {
                    // Access the cell value using the column name or index
                    object cellValue = row[column];

                    // Print the cell value on the console
                    Console.WriteLine($"Column: {column.ColumnName}, Value: {cellValue}");
                }
            }
        }

        // Iterate through each row and column in the worksheet
        // foreach (var row in worksheet.Rows)
        // {

        //     foreach (var cell in row.Cells)
        //     {
        //         // Get the cell value
        //         string cellValue = cell.Value?.ToString();
        //         if(cell.Row == 7)
        //         // Print the cell value on the console
        //             Console.WriteLine($"Cell [{cell.Row}, {cell.Column}]: {cellValue}");
        //     }
        // }

        // Dispose the workbook
        workbook.Dispose();
    }
}
