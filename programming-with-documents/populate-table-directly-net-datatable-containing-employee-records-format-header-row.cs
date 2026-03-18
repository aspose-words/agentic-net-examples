using System;
using System.Data;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;

class PopulateTableFromDataTable
{
    static void Main()
    {
        // Prepare a sample DataTable with employee data.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("EmployeeID", typeof(int));
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Department", typeof(string));
        employees.Columns.Add("Salary", typeof(decimal));

        employees.Rows.Add(1, "John Doe", "Sales", 55000m);
        employees.Rows.Add(2, "Jane Smith", "Marketing", 62000m);
        employees.Rows.Add(3, "Bob Johnson", "IT", 72000m);
        employees.Rows.Add(4, "Alice Brown", "HR", 48000m);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- Header Row ----------
        // Apply header formatting: bold text and light gray shading.
        builder.Font.Bold = true;
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.LightGray;

        // Insert header cells using column names from the DataTable.
        foreach (DataColumn col in employees.Columns)
        {
            builder.InsertCell();
            builder.Write(col.ColumnName);
        }
        // End the header row.
        builder.EndRow();

        // Reset formatting for data rows.
        builder.Font.Bold = false;
        builder.CellFormat.Shading.BackgroundPatternColor = System.Drawing.Color.Empty;

        // ---------- Data Rows ----------
        foreach (DataRow row in employees.Rows)
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

        // Optional: Auto‑fit the table to its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document.
        doc.Save("EmployeeReport.docx", SaveFormat.Docx);
    }
}
