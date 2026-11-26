using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table using the builder.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1,1");
        builder.InsertCell();
        builder.Writeln("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 2,1");
        builder.InsertCell();
        builder.Writeln("Cell 2,2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table from the document.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Configure each row to avoid breaking across pages.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.AllowBreakAcrossPages = false;
        }

        // Save the document to disk.
        string outputPath = "TableAllowBreakAcrossPages.docx";
        doc.Save(outputPath);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new Exception($"Failed to create the output file: {outputPath}");
        }
    }
}
