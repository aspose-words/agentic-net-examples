using System;
using System.Data;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Prepare a DataTable with employee information.
        DataTable employees = new DataTable("Employees");
        employees.Columns.Add("ID", typeof(int));
        employees.Columns.Add("Name", typeof(string));
        employees.Columns.Add("Department", typeof(string));
        employees.Columns.Add("Salary", typeof(decimal));

        employees.Rows.Add(1, "John Doe", "Finance", 60000m);
        employees.Rows.Add(2, "Jane Smith", "HR", 55000m);
        employees.Rows.Add(3, "Bob Johnson", "IT", 72000m);

        // Create a blank Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        // Apply shading and bold font to header cells.
        builder.RowFormat.Height = 20;
        builder.RowFormat.HeightRule = HeightRule.AtLeast;
        builder.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        builder.Font.Bold = true;

        InsertHeaderCell(builder, "ID");
        InsertHeaderCell(builder, "Name");
        InsertHeaderCell(builder, "Department");
        InsertHeaderCell(builder, "Salary");
        builder.EndRow();

        // ---------- Data rows ----------
        // Reset formatting for regular rows.
        builder.CellFormat.Shading.ClearFormatting();
        builder.Font.Bold = false;

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

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("EmployeeTable.docx");
    }

    // Helper method to insert a header cell with the given text.
    private static void InsertHeaderCell(DocumentBuilder builder, string text)
    {
        builder.InsertCell();
        builder.Write(text);
    }
}
