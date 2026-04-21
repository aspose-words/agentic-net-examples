using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a temporary PNG image file from an embedded base‑64 string.
        string imagePath = Path.Combine(Directory.GetCurrentDirectory(), "sample.png");
        CreateSampleImage(imagePath);

        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 1x1 table.
        builder.StartTable();
        builder.InsertCell(); // First (and only) cell.

        // Insert the image as a floating shape inside the cell.
        Shape pictureShape = builder.InsertImage(imagePath);
        pictureShape.WrapType = WrapType.None;          // Make it a floating shape.
        pictureShape.IsLayoutInCell = true;             // Ensure layout is inside the cell.

        // Validate that the property was set correctly.
        if (!pictureShape.IsLayoutInCell)
            throw new InvalidOperationException("IsLayoutInCell was not set to true.");

        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "PictureInTable.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);

        // Clean up the temporary image.
        File.Delete(imagePath);
    }

    // Writes a tiny PNG image (1×1 pixel) to the specified path.
    private static void CreateSampleImage(string path)
    {
        // Base‑64 encoded PNG (a single red pixel).
        const string base64Png =
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/5+BAQAE/wJ/lKXcAAAAAElFTkSuQmCC";

        byte[] imageBytes = Convert.FromBase64String(base64Png);
        File.WriteAllBytes(path, imageBytes);
    }
}
