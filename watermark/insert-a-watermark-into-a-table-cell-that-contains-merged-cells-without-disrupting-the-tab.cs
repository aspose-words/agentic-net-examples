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

        // Create a simple 1x1 pixel PNG (transparent) from a base64 string.
        const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8Xw8AAukB9YVYB2cAAAAASUVORK5CYII=";
        byte[] pngBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell in the first row.
        builder.StartTable();

        // First cell – start of merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Second cell – continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No text needed for the merged part.

        // Third cell – normal cell.
        builder.InsertCell();
        builder.Write("Normal Cell");

        builder.EndRow();

        // Second row – regular cells to keep table layout intact.
        builder.InsertCell();
        builder.Write("Row2 Cell1");
        builder.InsertCell();
        builder.Write("Row2 Cell2");
        builder.InsertCell();
        builder.Write("Row2 Cell3");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the merged cell (first cell of the first row).
        Table table = doc.FirstSection.Body.Tables[0];
        Cell mergedCell = table.Rows[0].Cells[0];

        // Move the builder cursor to the first paragraph of the merged cell.
        builder.MoveTo(mergedCell.FirstParagraph);

        // Insert the image as a shape inside the cell.
        Shape watermarkShape = builder.InsertImage(imagePath);
        // Configure the shape to act as a watermark (behind text, no wrapping).
        watermarkShape.WrapType = WrapType.None;
        watermarkShape.BehindText = true;
        // Set a semi‑transparent fill color (light gray) to simulate a watermark.
        watermarkShape.FillColor = System.Drawing.Color.FromArgb(50, System.Drawing.Color.LightGray);
        // Scale the image to fit the cell width while preserving aspect ratio.
        // CellFormat.Width is available; use it to size the shape.
        double cellWidth = mergedCell.CellFormat.Width;
        if (cellWidth > 0)
        {
            watermarkShape.Width = cellWidth - 10; // small margin
            watermarkShape.Height = watermarkShape.Width; // keep aspect ratio (square placeholder)
        }

        // Save the document.
        doc.Save(docPath);

        // Simple validation: ensure the output file exists.
        Console.WriteLine(File.Exists(docPath)
            ? $"Document created successfully: {docPath}"
            : "Failed to create the document.");
    }
}
