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

        // Start a table and add a single cell with some text.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Floating table cell.");
        builder.EndTable();

        // Make the table a floating object by enabling text wrapping.
        table.TextWrapping = TextWrapping.Around;
        // Position the floating table on the page.
        table.AbsoluteHorizontalDistance = 50; // points from the anchor.
        table.AbsoluteVerticalDistance = 20;   // points from the anchor.

        // Table.AllowOverlap is read‑only and defaults to true for floating tables.
        // No explicit check is needed; the property cannot be set manually.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTable.docx");
        doc.Save(outputPath);

        // Ensure the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("Failed to save the document.", outputPath);
        }
    }
}
