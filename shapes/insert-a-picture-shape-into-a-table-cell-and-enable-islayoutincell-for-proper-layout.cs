using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a temporary image file to insert.
        string imagePath = Path.Combine(Path.GetTempPath(), "sample_image.png");
        CreateSampleImage(imagePath);

        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with a single cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Set a fixed height for the row to make layout visible.
        builder.RowFormat.Height = 100;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Insert the picture shape into the current cell.
        Shape pictureShape = builder.InsertImage(imagePath);
        pictureShape.IsLayoutInCell = true; // Enable layout inside the cell.
        pictureShape.Width = 80;
        pictureShape.Height = 80;

        // Finish the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Path.GetTempPath(), "Output.docx");
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");

        // Optional cleanup of the temporary image.
        // File.Delete(imagePath);
    }

    // Creates a minimal 1x1 pixel PNG image and writes it to the specified path.
    private static void CreateSampleImage(string path)
    {
        // PNG data for a 1x1 transparent pixel.
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
