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

        // First cell (row 0, column 0)
        builder.InsertCell();
        builder.Write("Cell 1,1");

        // Second cell (row 0, column 1) – this is the target cell.
        builder.InsertCell();
        builder.Write("Target Cell");

        // End first row.
        builder.EndRow();

        // Third cell (row 1, column 0)
        builder.InsertCell();
        builder.Write("Cell 2,1");

        // Fourth cell (row 1, column 1)
        builder.InsertCell();
        builder.Write("Cell 2,2");

        // End the table.
        builder.EndTable();

        // Apply a text watermark to the document.
        // The watermark will be visible behind all content, including the target cell.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "WatermarkedTable.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document saved successfully: " + outputPath);
        }
        else
        {
            Console.WriteLine("Failed to save the document.");
        }
    }
}
