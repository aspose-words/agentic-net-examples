using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a DataTable with employee information.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("ID", typeof(int));
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Department", typeof(string));
        employees.Columns.Add("Salary", typeof(decimal));

        employees.Rows.Add(1, "Alice Johnson", "HR", 55000m);
        employees.Rows.Add(2, "Bob Smith", "IT", 72000m);
        employees.Rows.Add(3, "Carol White", "Finance", 68000m);

        // Create a new blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        Table table = builder.StartTable();

        // ----- Header row formatting -----
        builder.Font.Bold = true;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Insert header cells.
        InsertHeaderCell(builder, "ID");
        InsertHeaderCell(builder, "Name");
        InsertHeaderCell(builder, "Department");
        InsertHeaderCell(builder, "Salary");
        builder.EndRow();

        // ----- Data rows formatting -----
        builder.Font.Bold = false;
        builder.CellFormat.Shading.ClearFormatting();
        builder.ParagraphFormat.ClearFormatting();

        // Populate the table with data from the DataTable.
        foreach (DataRow row in employees.Rows)
        {
            builder.InsertCell();
            builder.Write(row["ID"].ToString());

            builder.InsertCell();
            builder.Write(row["Name"].ToString());

            builder.InsertCell();
            builder.Write(row["Department"].ToString());

            builder.InsertCell();
            builder.Write(string.Format("{0:C}", row["Salary"]));

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "EmployeeTable.docx";
        doc.Save(outputPath);
    }

    // Helper method to insert a header cell with the current formatting.
    private static void InsertHeaderCell(DocumentBuilder builder, string text)
    {
        builder.InsertCell();
        builder.Write(text);
    }
}
