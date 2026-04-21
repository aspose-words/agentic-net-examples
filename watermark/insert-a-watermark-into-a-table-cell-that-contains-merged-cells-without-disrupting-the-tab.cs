using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directories
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string imagePath = Path.Combine(artifactsDir, "watermark.png");

        // Create a simple 1x1 transparent PNG (base64 encoded) and save it locally
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Create a new document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell (spanning two columns)
        builder.StartTable();

        // First cell – start of merged range
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell Content");

        // Second cell – merged with the first
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part

        builder.EndRow();

        // Add a normal row below to keep table layout intact
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Move cursor to the merged cell (first row, first column)
        builder.MoveToCell(0, 0, 0, 0);

        // Insert the image as a watermark inside the cell
        Shape watermarkShape = builder.InsertImage(imagePath);
        watermarkShape.WrapType = WrapType.None;          // No text wrapping
        watermarkShape.BehindText = true;                // Place behind cell text
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
        // Scale the image to fit the cell (optional)
        watermarkShape.Width = builder.CellFormat.Width;
        watermarkShape.Height = builder.RowFormat.Height > 0 ? builder.RowFormat.Height : 50;

        // Save the document
        string outputPath = Path.Combine(artifactsDir, "TableCellWatermark.docx");
        doc.Save(outputPath);

        // Simple validation that the file was created
        if (File.Exists(outputPath))
        {
            Console.WriteLine("Document created successfully: " + outputPath);
        }
    }
}
