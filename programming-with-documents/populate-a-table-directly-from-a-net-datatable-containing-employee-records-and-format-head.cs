using System;
using System.Data;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

namespace AsposeWordsTableFromDataTable
{
    public class Program
    {
        public static void Main()
        {
            // Create a DataTable with employee information.
            DataTable employeeTable = new DataTable("Employees");
            employeeTable.Columns.Add("ID", typeof(int));
            employeeTable.Columns.Add("Name", typeof(string));
            employeeTable.Columns.Add("Department", typeof(string));
            employeeTable.Columns.Add("Salary", typeof(decimal));

            // Add sample rows.
            employeeTable.Rows.Add(1, "John Doe", "HR", 50000m);
            employeeTable.Rows.Add(2, "Jane Smith", "IT", 65000m);
            employeeTable.Rows.Add(3, "Bob Johnson", "Finance", 72000m);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start a new table.
            Table table = builder.StartTable();

            // ----- Header row -----
            // Make the header row repeat on each page (optional).
            builder.RowFormat.HeadingFormat = true;
            // Apply bold font and a light gray background to header cells.
            builder.Font.Bold = true;
            builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

            // Insert header cells using column names from the DataTable.
            foreach (DataColumn column in employeeTable.Columns)
            {
                builder.InsertCell();
                builder.Write(column.ColumnName);
            }
            builder.EndRow();

            // ----- Data rows -----
            // Reset formatting for regular rows.
            builder.Font.Bold = false;
            builder.CellFormat.Shading.ClearFormatting();

            // Populate the table with the DataTable rows.
            foreach (DataRow dataRow in employeeTable.Rows)
            {
                foreach (object cellValue in dataRow.ItemArray)
                {
                    builder.InsertCell();
                    builder.Write(cellValue?.ToString() ?? string.Empty);
                }
                builder.EndRow();
            }

            // Finish the table.
            builder.EndTable();

            // Adjust column widths to fit the content.
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document to the current working directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "EmployeeTable.docx");
            doc.Save(outputPath);
        }
    }
}
