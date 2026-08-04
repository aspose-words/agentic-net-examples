using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file names
        const string sampleImagePath = "sample.png";
        const string placeholderImagePath = "placeholder.png";
        const string originalDocPath = "original.docx";
        const string modifiedDocPath = "modified.docx";

        // Create a sample image to be used in the document
        CreateSampleImage(sampleImagePath, 200, 100, Aspose.Drawing.Color.LightBlue);

        // Create a placeholder image that will replace existing images
        CreateSampleImage(placeholderImagePath, 200, 100, Aspose.Drawing.Color.LightGray);

        // Build a document that contains a few images
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Document with images:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        builder.Writeln();
        builder.InsertImage(sampleImagePath);
        doc.Save(originalDocPath);

        // Load the document we just created
        var loadedDoc = new Document(originalDocPath);

        // Replace each image with the placeholder image
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Replace the image data
                shape.ImageData.SetImage(placeholderImagePath);
            }
        }

        // Save the modified document
        loadedDoc.Save(modifiedDocPath);

        // Simple validation
        if (!File.Exists(modifiedDocPath))
            throw new InvalidOperationException("Modified document was not saved.");

        // Clean up generated files (optional)
        // File.Delete(sampleImagePath);
        // File.Delete(placeholderImagePath);
        // File.Delete(originalDocPath);
        // File.Delete(modifiedDocPath);
    }

    // Helper method to create a solid‑color bitmap and save it to a file
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color color)
    {
        using (var bitmap = new Bitmap(width, height))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(color);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
