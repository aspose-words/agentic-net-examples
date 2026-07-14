using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

namespace WatermarkInTableCell
{
    public class Program
    {
        public static void Main()
        {
            // Define paths.
            string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
            Directory.CreateDirectory(outputDir);
            string imagePath = Path.Combine(outputDir, "watermark.png");
            string docPath = Path.Combine(outputDir, "TableCellWatermark.docx");

            // Create a simple PNG image (red square) from a Base64 string.
            const string base64Png =
                "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAYAAABw4pVUAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJ" +
                "bWFnZVJlYWR5ccllPAAAABh0RVh0U291cmNlAGh0dHA6Ly93d3cuaW1hZ2UuanBnAAAAAElFTkSuQmCC";
            byte[] pngBytes = Convert.FromBase64String(base64Png);
            File.WriteAllBytes(imagePath, pngBytes);

            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Build a 4x4 table.
            Table table = builder.StartTable();

            // First row.
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R1C{col + 1}");
            }
            builder.EndRow();

            // Second row.
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R2C{col + 1}");
            }
            builder.EndRow();

            // Third row.
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R3C{col + 1}");
            }
            builder.EndRow();

            // Fourth row.
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R4C{col + 1}");
            }
            builder.EndRow();

            builder.EndTable();

            // Merge cells to create a spanning cell (rows 2-3, columns 2-3).
            // Horizontal merge.
            builder.MoveToCell(0, 0, 1, 1); // Row 2, Column 2 (0‑based indices).
            builder.CellFormat.HorizontalMerge = CellMerge.First;
            builder.Write("Merged Cell");

            builder.MoveToCell(0, 0, 1, 2); // Row 2, Column 3.
            builder.CellFormat.HorizontalMerge = CellMerge.Previous;

            // Vertical merge for the same columns.
            builder.MoveToCell(0, 0, 2, 1); // Row 3, Column 2.
            builder.CellFormat.VerticalMerge = CellMerge.First;

            builder.MoveToCell(0, 0, 3, 1); // Row 4, Column 2 (will not be merged, just to reset later).
            builder.CellFormat.VerticalMerge = CellMerge.None;

            builder.MoveToCell(0, 0, 2, 2); // Row 3, Column 3.
            builder.CellFormat.VerticalMerge = CellMerge.Previous;

            // Reset merge settings for subsequent cells.
            builder.CellFormat.HorizontalMerge = CellMerge.None;
            builder.CellFormat.VerticalMerge = CellMerge.None;

            // Move cursor to the top‑left cell of the merged region to insert the watermark image.
            builder.MoveToCell(0, 0, 1, 1);
            Shape watermarkShape = builder.InsertImage(imagePath);
            watermarkShape.WrapType = WrapType.None;
            watermarkShape.BehindText = true;
            watermarkShape.RelativeHorizontalPosition = RelativeHorizontalPosition.Column;
            watermarkShape.RelativeVerticalPosition = RelativeVerticalPosition.Paragraph;
            // Scale the image to fit the merged cell (optional).
            watermarkShape.Width = builder.RowFormat.Height; // Example scaling; adjust as needed.
            watermarkShape.Height = builder.RowFormat.Height;

            // Save the document.
            doc.Save(docPath);

            // Simple validation: ensure the file was created.
            if (File.Exists(docPath))
            {
                Console.WriteLine($"Document saved successfully to: {docPath}");
            }
        }
    }
}
