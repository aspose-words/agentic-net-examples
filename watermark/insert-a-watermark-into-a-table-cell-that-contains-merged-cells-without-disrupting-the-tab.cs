using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output directories.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);
        string imagePath = Path.Combine(outputDir, "watermark.png");
        string docPath = Path.Combine(outputDir, "TableWithCellWatermark.docx");

        // Create a simple 1x1 pixel PNG image (red) from a Base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9W6XcVYAAAAASUVORK5CYII=";
        File.WriteAllBytes(imagePath, Convert.FromBase64String(base64Png));

        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell in the first row.
        builder.StartTable();

        // First cell – start of merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell Content");

        // Second cell – merged with the first.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        builder.EndRow();

        // Add a normal second row to keep the table layout intact.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.Write("Row 2, Cell 2");

        builder.EndRow();
        builder.EndTable();

        // Move the cursor to the first cell (the merged cell) to insert the watermark.
        Table table = doc.FirstSection.Body.Tables[0];
        Cell mergedCell = table.Rows[0].FirstCell;
        builder.MoveTo(mergedCell.FirstParagraph);

        // Insert the image as a floating shape so it can act as a watermark.
        Shape watermarkShape = builder.InsertImage(imagePath);
        watermarkShape.WrapType = WrapType.None;          // No text wrapping.
        watermarkShape.BehindText = true;                // Place behind the cell text.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

        // Scale the watermark to fit the cell width. Height is left to auto‑scale to preserve aspect ratio.
        watermarkShape.Width = mergedCell.CellFormat.Width;

        // Save the document.
        doc.Save(docPath);

        // Simple validation that the output file exists.
        if (File.Exists(docPath))
        {
            Console.WriteLine("Document created successfully: " + docPath);
        }
    }
}
