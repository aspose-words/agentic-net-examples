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

        // Build a 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");

        builder.InsertCell();
        builder.Writeln("Cell 2");

        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");

        builder.InsertCell();
        builder.Writeln("Cell 4");

        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the built‑in "Table Grid" style to the table by name.
        // The TableStyleIdentifier enum may not be available in all versions,
        // so we set the style using its exact name.
        table.StyleName = "Table Grid";

        // Save the document to a file.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException($"Failed to create the output file: {outputPath}");

        Console.WriteLine($"Document saved successfully to '{outputPath}'.");
    }
}
