using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class InsertPictureInTableCell
{
    public static void Main()
    {
        // Prepare a temporary folder for sample files.
        string tempFolder = Path.Combine(Path.GetTempPath(), "AsposeExample");
        Directory.CreateDirectory(tempFolder);

        // Create a simple placeholder PNG image (1x1 pixel, light blue background).
        string imagePath = Path.Combine(tempFolder, "sample.png");
        byte[] pngBytes = new byte[]
        {
            // PNG header
            0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
            // IHDR chunk (1x1 pixel, 8-bit RGBA)
            0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
            0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
            0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
            0x89,
            // IDAT chunk (pixel data)
            0x00,0x00,0x00,0x0A,0x49,0x44,0x41,0x54,
            0x78,0x9C,0x63,0x60,0x60,0x60,0x00,0x00,
            0x00,0x05,0x00,0x01,0xE2,0x26,0x05,0x9B,
            // IEND chunk
            0x00,0x00,0x00,0x00,0x49,0x45,0x4E,0x44,
            0xAE,0x42,0x60,0x82
        };
        File.WriteAllBytes(imagePath, pngBytes);

        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Insert the image into the current cell.
        builder.InsertImage(imagePath);

        // Retrieve the inserted shape (image) from the cell.
        Shape pictureShape = null;
        NodeCollection shapes = builder.CurrentParagraph.GetChildNodes(NodeType.Shape, false);
        foreach (Node node in shapes)
        {
            if (node is Shape shape && shape.HasImage)
            {
                pictureShape = shape;
                break;
            }
        }

        if (pictureShape == null)
            throw new InvalidOperationException("Image shape was not inserted.");

        // Enable layout inside the table cell.
        pictureShape.IsLayoutInCell = true;

        // End the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(tempFolder, "Output.docx");
        doc.Save(outputPath);

        // Validate that the document was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);

        // Validate that the shape property is set correctly.
        if (!pictureShape.IsLayoutInCell)
            throw new InvalidOperationException("IsLayoutInCell property was not set correctly.");
    }
}
