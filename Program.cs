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

        Workbook outputWorkbook = new Workbook();



        DataTable dt = worksheet.ExportDataTable(worksheet.AllocatedRange, false, true);
        Worksheet outputWorksheet = outputWorkbook.Worksheets[0];
        outputWorksheet.InsertDataTable(dt, true, 1, 1);
        outputWorkbook.SaveToFile(@"C:\Users\Honey.Kumar\OneDrive - Shell\Documents\Honey Kumar\workspace\blank3u.xlsx", FileFormat.Version2013);
        // int count= 0;
        // foreach (DataRow row in dt.Rows)
        // {
        //     count++;
        //     if (count == 6)
        //     {
        //     // Loop through each column in the DataRow
        //         foreach (DataColumn column in dt.Columns)
        //         {
        //             // Access the cell value using the column name or index
        //             object cellValue = row[column];

        //             // Print the cell value on the console
        //             Console.WriteLine($"Column: {column.ColumnName}, Value: {cellValue}");
        //         }
        //     }
        // }

     

        // Dispose the workbook
        workbook.Dispose();
        outputWorkbook.Dispose();
    }
}
