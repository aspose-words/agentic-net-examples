using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class InsertPictureInTableCell
{
    static void Main()
    {
        // Create a temporary 1x1 PNG image (if it does not already exist).
        string imagePath = Path.Combine(Path.GetTempPath(), "sample.png");
        if (!File.Exists(imagePath))
        {
            // Minimal PNG file (1×1 pixel, transparent).
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
            File.WriteAllBytes(imagePath, pngBytes);
        }

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1‑row, 2‑cell table.
        builder.StartTable();
        builder.InsertCell(); // First cell.
        builder.InsertCell(); // Second cell.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Move the cursor to the first paragraph of the first cell.
        builder.MoveTo(table.FirstRow.FirstCell.FirstParagraph);

        // Insert a floating image shape into the cell.
        Shape picture = builder.InsertImage(
            imagePath,
            RelativeHorizontalPosition.LeftMargin, 50,
            RelativeVerticalPosition.TopMargin, 100,
            100, // width in points
            100, // height in points
            WrapType.None);

        // Enable layout inside the cell so the shape moves with cell resizing.
        picture.IsLayoutInCell = true;

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the resulting document.
        string outputPath = Path.Combine(outputDir, "PictureInTableCell.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
