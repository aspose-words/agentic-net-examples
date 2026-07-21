using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Add a text watermark that will appear behind the document content.
        doc.Watermark.SetText("Cell Watermark");

        // Use DocumentBuilder to construct a table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Define the number of cells in the first row.
        int firstRowCellCount = 3;

        // Populate the first row with cells.
        for (int i = 0; i < firstRowCellCount; i++)
        {
            // Insert a new cell.
            builder.InsertCell();

            // Write some sample text inside the cell.
            builder.Write($"Cell {i + 1}");
        }

        // End the first row.
        builder.EndRow();

        // Add a second row to demonstrate normal table layout.
        for (int i = 0; i < firstRowCellCount; i++)
        {
            builder.InsertCell();
            builder.Write($"Row2 Cell {i + 1}");
        }
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "TableWithCellWatermark.docx");
        doc.Save(outputPath);
    }
}
