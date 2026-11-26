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

        // Build a simple 2x2 table.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Retrieve the created table (the first table in the document).
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Apply the built‑in TableGrid style using the StyleIdentifier enum.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Save the document to a file.
        string outputPath = "TableStyleExample.docx";
        doc.Save(outputPath);

        // Verify that the file was saved successfully.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
