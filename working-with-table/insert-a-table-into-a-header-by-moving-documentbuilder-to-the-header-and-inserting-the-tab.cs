using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder cursor to the primary header of the first section.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);

        // Start a table inside the header.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 1");

        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 2");
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Header Cell 3");

        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Header Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Return the cursor to the main document body.
        builder.MoveToSection(0);

        // Add a simple paragraph to the body so the document has visible content.
        builder.Writeln("Document body content.");

        // Define the output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "HeaderTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");

        // Optionally, inform that the process completed (no interactive prompts required).
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
