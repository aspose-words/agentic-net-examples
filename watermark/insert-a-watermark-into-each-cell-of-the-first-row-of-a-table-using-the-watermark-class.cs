using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class WatermarkInTableCells
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // First row – three cells.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Header {i + 1}");
        }
        // End the first row.
        builder.EndRow();

        // Second row – sample data.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Write($"Row 1, Cell {i + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Add a text watermark to the document.
        // The watermark will be visible behind the content of every cell,
        // effectively appearing in each cell of the first row.
        doc.Watermark.SetText("CONFIDENTIAL");

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithWatermark.docx");
        doc.Save(outputPath);
    }
}
