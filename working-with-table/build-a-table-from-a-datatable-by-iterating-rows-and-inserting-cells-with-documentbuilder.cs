using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample DataTable.
        DataTable dataTable = new DataTable("Sample");
        dataTable.Columns.Add("Product");
        dataTable.Columns.Add("Quantity", typeof(int));
        dataTable.Columns.Add("Price", typeof(decimal));

        dataTable.Rows.Add("Apples", 10, 0.5m);
        dataTable.Rows.Add("Bananas", 5, 0.3m);
        dataTable.Rows.Add("Carrots", 7, 0.2m);

        // Initialize a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        Table table = builder.StartTable();

        // Insert header cells.
        foreach (DataColumn column in dataTable.Columns)
        {
            builder.InsertCell();
            builder.Write(column.ColumnName);
        }
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object value in row.ItemArray)
            {
                builder.InsertCell();
                builder.Write(value?.ToString() ?? string.Empty);
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableFromDataTable.docx");
        doc.Save(outputPath);
    }
}
