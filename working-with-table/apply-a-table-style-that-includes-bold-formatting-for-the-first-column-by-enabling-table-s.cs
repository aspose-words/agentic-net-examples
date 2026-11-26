using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 3x2 table.
        builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Row 1, Col 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Col 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Writeln("Row 2, Col 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Col 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply a built‑in style.
        table.Style = doc.Styles["Table Grid"];

        // Enable bold formatting for the first column using the supported option.
        table.StyleOptions = TableStyleOptions.FirstColumn;

        // Save the document.
        string outputPath = "TableFirstColumnBold.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        Console.WriteLine($"Document saved successfully to '{Path.GetFullPath(outputPath)}'.");
    }
}
