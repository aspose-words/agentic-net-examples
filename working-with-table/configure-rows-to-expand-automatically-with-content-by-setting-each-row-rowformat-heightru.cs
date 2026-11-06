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

        // Start a new table.
        builder.StartTable();

        // First row with short content.
        builder.InsertCell();
        builder.Writeln("Short");
        builder.EndRow();

        // Second row with longer content to demonstrate auto height.
        builder.InsertCell();
        builder.Writeln("This is a longer piece of text that should cause the row to expand automatically based on its content.");
        builder.EndRow();

        // Third row with even longer content.
        builder.InsertCell();
        builder.Writeln("Another row with a very long text. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set each row's HeightRule to Auto so rows expand with their content.
        foreach (Row row in table.Rows)
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document to a file.
        string outputPath = "TableRowsAutoHeight.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");
        }

        // Optionally, inform that the process completed (no interactive input required).
        Console.WriteLine("Document created successfully: " + Path.GetFullPath(outputPath));
    }
}
