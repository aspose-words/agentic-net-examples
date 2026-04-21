using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x4 table.
        Table table = builder.StartTable();

        // First cell will span 2 columns and 2 rows.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Merged Cell");

        // Second cell in the first row (merged horizontally).
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write(string.Empty); // placeholder

        builder.EndRow();

        // Second row – cells that continue the vertical merge.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.Write(string.Empty); // placeholder

        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Cell 2,2");

        builder.EndRow();

        // Add two more rows with regular cells.
        for (int i = 0; i < 2; i++)
        {
            builder.InsertCell();
            builder.Write($"R{i + 3}C1");
            builder.InsertCell();
            builder.Write($"R{i + 3}C2");
            builder.EndRow();
        }

        builder.EndTable();

        // Retrieve the merged cell (first row, first cell).
        Cell mergedCell = table.Rows[0].Cells[0];

        // Move the builder cursor to the first paragraph of the merged cell.
        builder.MoveTo(mergedCell.FirstParagraph);

        // A simple 1x1 transparent PNG (base64 encoded).
        string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
        byte[] imageBytes = Convert.FromBase64String(base64Png);

        // Insert the image as a floating shape.
        Shape watermarkShape = builder.InsertImage(imageBytes);
        watermarkShape.WrapType = WrapType.None;
        watermarkShape.BehindText = true;

        // Position the shape relative to the column (the cell) and paragraph.
        watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
        watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;

        // Resize the shape to cover the cell.
        // Width comes from the cell's format.
        watermarkShape.Width = mergedCell.CellFormat.Width;

        // Height: use the height of the first row that the merged cell belongs to.
        // If the height is not set, fall back to a reasonable default.
        double rowHeight = mergedCell.ParentRow.RowFormat.Height;
        watermarkShape.Height = rowHeight > 0 ? rowHeight : 50;

        // Align shape to the top‑left corner of the cell.
        watermarkShape.Left = 0;
        watermarkShape.Top = 0;

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellWatermark.docx");
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
