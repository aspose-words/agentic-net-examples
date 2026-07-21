using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a simple 1x1 PNG file that will be used as the watermark image.
        string imagePath = "watermark.png";
        CreateSamplePng(imagePath);

        // Create a blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a 4x4 table with sample text in each cell.
        Table table = builder.StartTable();
        for (int row = 0; row < 4; row++)
        {
            for (int col = 0; col < 4; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // Merge cells to create a cell that spans rows 1‑2 and columns 1‑2.
        // Top‑left cell of the span.
        builder.MoveToCell(0, 0, 0, 0);
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.CellFormat.VerticalMerge = CellMerge.First;

        // Cells that join the span horizontally.
        builder.MoveToCell(0, 0, 0, 1);
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // Cells that join the span vertically.
        builder.MoveToCell(0, 1, 0, 0);
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Insert the image watermark into the merged cell.
        builder.MoveToCell(0, 0, 0, 0);
        builder.InsertImage(imagePath);

        // Retrieve the inserted shape (the image) and configure it as a watermark.
        Shape shape = (Shape)builder.CurrentParagraph.GetChildNodes(NodeType.Shape, false)[0];
        shape.WrapType = WrapType.None;   // No text wrapping.
        shape.BehindText = true;          // Place behind the cell text.

        // Save the document.
        doc.Save("WatermarkedTable.docx");
    }

    // Writes a minimal 1x1 pixel PNG to the specified path.
    private static void CreateSamplePng(string path)
    {
        // PNG data for a single transparent pixel.
        byte[] pngBytes = new byte[]
        {
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
            0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
            0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
            0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
            0x42,0x60,0x82
        };
        File.WriteAllBytes(path, pngBytes);
    }
}
