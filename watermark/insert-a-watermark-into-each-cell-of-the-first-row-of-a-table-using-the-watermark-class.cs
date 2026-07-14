using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Add a text watermark that will appear behind the document content,
        // thus it will be visible in every cell of the table.
        doc.Watermark.SetText("Cell Watermark");

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with three columns.
        builder.StartTable();

        // Populate the first row with three cells.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row 1, Cell {i + 1}");
        }

        // End the first row.
        builder.EndRow();

        // Populate a second row to demonstrate normal table layout.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row 2, Cell {i + 1}");
        }

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine($"Document saved successfully to: {outputPath}");
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
