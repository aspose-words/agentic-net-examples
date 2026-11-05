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
        dataTable.Columns.Add("ID", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Score", typeof(double));

        dataTable.Rows.Add(1, "Alice", 85.5);
        dataTable.Rows.Add(2, "Bob", 92.0);
        dataTable.Rows.Add(3, "Charlie", 78.3);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building the table.
        builder.StartTable();

        // Insert header row.
        foreach (DataColumn column in dataTable.Columns)
        {
            builder.InsertCell();
            builder.Write(column.ColumnName);
        }
        builder.EndRow();

        // Insert data rows.
        foreach (DataRow row in dataTable.Rows)
        {
            foreach (DataColumn column in dataTable.Columns)
            {
                builder.InsertCell();
                builder.Write(row[column].ToString());
            }
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = "TableFromDataTable.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception($"Failed to create the output file: {outputPath}");

        // Re-open the document to verify table structure.
        Document loadedDoc = new Document(outputPath);
        NodeCollection tables = loadedDoc.GetChildNodes(NodeType.Table, true);
        if (tables.Count == 0)
            throw new Exception("No tables found in the saved document.");

        Table table = (Table)tables[0];
        // Expected rows: header + data rows.
        int expectedRowCount = dataTable.Rows.Count + 1;
        if (table.Rows.Count != expectedRowCount)
            throw new Exception($"Table row count mismatch. Expected: {expectedRowCount}, Actual: {table.Rows.Count}");

        // If execution reaches this point, the table was built successfully.
        Console.WriteLine("Table created and saved successfully.");
    }
}
