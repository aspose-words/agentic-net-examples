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

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Apply the built‑in TableGrid style using the style identifier.
        // Using StyleIdentifier avoids a null lookup in the Styles collection.
        table.StyleIdentifier = StyleIdentifier.TableGrid;
        // Optional: apply the style to the whole table.
        table.StyleOptions = TableStyleOptions.None;

        // Save the document to a local file.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(outputDir);
        string outputPath = Path.Combine(outputDir, "TableWithStyle.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved correctly.");

        // Indicate successful completion.
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
