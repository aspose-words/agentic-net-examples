using System;
using System.Data;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample DataTable with some data.
        DataTable dataTable = new DataTable("Sample");
        dataTable.Columns.Add("Product");
        dataTable.Columns.Add("Quantity");
        dataTable.Columns.Add("Price");
        dataTable.Rows.Add("Apples", 10, 1.5);
        dataTable.Rows.Add("Bananas", 5, 0.8);
        dataTable.Rows.Add("Carrots", 7, 0.6);

        // Create a new blank document and a builder attached to it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Add a header row using the column names.
        foreach (DataColumn column in dataTable.Columns)
        {
            builder.InsertCell();
            builder.Write(column.ColumnName);
        }
        builder.EndRow();

        // Iterate through each DataRow and add its values to the table.
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

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file at {outputPath}");
        }
    }
}
