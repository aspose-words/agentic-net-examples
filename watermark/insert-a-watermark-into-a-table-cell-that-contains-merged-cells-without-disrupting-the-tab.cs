using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class TableCellWatermarkExample
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // Create a simple 1x1 transparent PNG to act as the watermark.
        string imagePath = Path.Combine(artifactsDir, "watermark.png");
        byte[] pngBytes = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK6cAAAAASUVORK5CYII=");
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with a horizontally merged cell in the first row.
        Table table = builder.StartTable();

        // First cell – start of the merged range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Merged cell content");

        // Second cell – continues the merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // Second row – normal unmerged cells.
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Retrieve the merged cell (first cell of the first row).
        Cell mergedCell = table.FirstRow.FirstCell;

        // Move the builder cursor into the merged cell.
        builder.MoveTo(mergedCell.FirstParagraph);

        // Insert the image as a floating shape inside the cell.
        Shape watermarkShape = builder.InsertImage(imagePath);
        watermarkShape.WrapType = WrapType.None;          // No text wrapping.
        watermarkShape.BehindText = true;                // Place behind cell text.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column; // Position relative to the column.

        // Adjust size to fill the cell.
        // Width is taken from the cell's format; height from the parent row's format.
        watermarkShape.Width = mergedCell.CellFormat.Width;
        watermarkShape.Height = mergedCell.ParentRow.RowFormat.Height;

        // Save the document.
        string outputPath = Path.Combine(artifactsDir, "TableCellWatermark.docx");
        doc.Save(outputPath);
    }
}
