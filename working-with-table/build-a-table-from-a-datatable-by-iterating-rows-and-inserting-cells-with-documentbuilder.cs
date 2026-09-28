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
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Score", typeof(double));

        dataTable.Rows.Add(1, "Alice", 85.5);
        dataTable.Rows.Add(2, "Bob", 92.0);
        dataTable.Rows.Add(3, "Charlie", 78.0);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building the table.
        builder.StartTable();

        // Insert header row.
        foreach (DataColumn column in dataTable.Columns)
        {
            builder.InsertCell();
            builder.Writeln(column.ColumnName);
        }
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (object value in row.ItemArray)
            {
                builder.InsertCell();
                builder.Writeln(value?.ToString() ?? string.Empty);
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        string outputPath = "TableFromDataTable.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
    }
}
