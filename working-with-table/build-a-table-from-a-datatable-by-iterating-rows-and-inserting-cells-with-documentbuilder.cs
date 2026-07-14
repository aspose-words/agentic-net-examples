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
        dataTable.Rows.Add("Apples", 10);
        dataTable.Rows.Add("Bananas", 20);
        dataTable.Rows.Add("Carrots", 30);

        // Initialize a new blank document and a builder.
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

        // Add rows for each DataRow in the DataTable.
        foreach (DataRow dataRow in dataTable.Rows)
        {
            foreach (object value in dataRow.ItemArray)
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

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The output document was not created.");
        }
    }
}
