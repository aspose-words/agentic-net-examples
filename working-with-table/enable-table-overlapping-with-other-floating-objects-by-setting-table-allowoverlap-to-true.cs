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

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a single cell with some text.
        builder.InsertCell();
        builder.Write("Floating table cell.");

        // Finish the table.
        builder.EndTable();

        // Make the table a floating object by enabling text wrapping.
        table.TextWrapping = TextWrapping.Around;

        // Position the floating table (optional – just for demonstration).
        table.AbsoluteHorizontalDistance = 50; // 50 points from the paragraph.
        table.AbsoluteVerticalDistance = 20;   // 20 points from the paragraph.

        // The AllowOverlap property is read‑only and defaults to true.
        // No need to set it; just ensure the document is saved.
        if (!table.AllowOverlap)
        {
            // If for some reason overlapping is disabled, we could log or handle it here.
            // Throwing an exception is not required for this task.
            Console.WriteLine("Warning: Table overlapping is disabled.");
        }

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "FloatingTableAllowOverlap.docx");

        // Save the document.
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
        {
            throw new FileNotFoundException("The output document was not saved correctly.", outputPath);
        }

        // Inform that the process completed successfully.
        Console.WriteLine("Document saved to: " + outputPath);
    }
}
