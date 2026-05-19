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

        // Make the table floating and wrap text around it.
        table.TextWrapping = TextWrapping.Around;
        table.AbsoluteHorizontalDistance = 50; // Horizontal offset in points.
        table.AbsoluteVerticalDistance = 20;   // Vertical offset in points.

        // The AllowOverlap property is read‑only and defaults to true for newly created tables.
        // No additional action is required to enable overlapping.

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableAllowOverlap.docx");
        doc.Save(outputPath);

        // Ensure the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("Failed to save the document.", outputPath);
    }
}
