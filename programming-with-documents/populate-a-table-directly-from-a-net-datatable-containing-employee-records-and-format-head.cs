using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with employee data.
        DataTable employeeTable = new DataTable("Employees");
        employeeTable.Columns.Add("ID", typeof(int));
        employeeTable.Columns.Add("Name", typeof(string));
        employeeTable.Columns.Add("Department", typeof(string));
        employeeTable.Columns.Add("Salary", typeof(decimal));

        employeeTable.Rows.Add(1, "Alice Johnson", "Finance", 75000m);
        employeeTable.Rows.Add(2, "Bob Smith", "Engineering", 82000m);
        employeeTable.Rows.Add(3, "Carol White", "HR", 63000m);
        employeeTable.Rows.Add(4, "David Brown", "Marketing", 58000m);

        // Create a new Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table.
        Table table = builder.StartTable();

        // Define header formatting.
        builder.Font.Bold = true;
        builder.Font.Size = 12;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;

        // Populate header row.
        foreach (DataColumn column in employeeTable.Columns)
        {
            builder.InsertCell();
            builder.Write(column.ColumnName);
        }
        builder.EndRow();

        // Reset formatting for data rows.
        builder.Font.Bold = false;
        builder.Font.Size = 10;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.Empty;

        // Populate data rows.
        foreach (DataRow row in employeeTable.Rows)
        {
            foreach (object cellValue in row.ItemArray)
            {
                builder.InsertCell();
                builder.Write(cellValue?.ToString() ?? string.Empty);
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "EmployeeReport.docx";
        doc.Save(outputPath);
    }
}
