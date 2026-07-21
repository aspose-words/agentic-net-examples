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

        // Use DocumentBuilder to add a table with two rows and two columns.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // First cell – this is where we want the watermark to be visible.
        builder.InsertCell();
        builder.Write("Cell with watermark");

        // Second cell in the first row.
        builder.InsertCell();
        builder.Write("Another cell");

        // End the first row.
        builder.EndRow();

        // Second row cells.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");

        // End the table.
        builder.EndTable();

        // Add a text watermark to the document.
        // The watermark will appear behind all page content, including the target cell.
        TextWatermarkOptions options = new TextWatermarkOptions
        {
            FontFamily = "Arial",
            FontSize = 36,
            Color = System.Drawing.Color.LightGray,
            Layout = WatermarkLayout.Diagonal,
            IsSemitrasparent = true
        };
        doc.Watermark.SetText("CONFIDENTIAL", options);

        // Define output path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellWatermark.docx");

        // Save the document.
        doc.Save(outputPath);

        // Simple validation: ensure the file was created.
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
