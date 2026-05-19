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

        // Start a table and add a simple 2x2 grid.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndTable(); // Ends the table.

        // Set spacing before and after the table (in points) using the correct properties.
        table.DistanceTop = 12;    // Space above the table.
        table.DistanceBottom = 12; // Space below the table.

        // Define an output path relative to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "TableSpacing.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
