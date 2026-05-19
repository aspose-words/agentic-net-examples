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
        // Define file names in the current directory.
        string originalDoc = "original.docx";
        string modifiedDoc = "modified.docx";
        string sampleImage = "sample.png";
        string placeholderImage = "placeholder.png";

        // -----------------------------------------------------------------
        // 1. Create a sample image that will be inserted into the document.
        // -----------------------------------------------------------------
        CreateSampleImage(sampleImage, 200, 200, Aspose.Drawing.Color.LightBlue);

        // -----------------------------------------------------------------
        // 2. Build a document and insert a few images.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three sample images separated by paragraphs.
        for (int i = 0; i < 3; i++)
        {
            builder.Writeln($"Paragraph before image {i + 1}:");
            builder.InsertImage(sampleImage);
            builder.Writeln($"Paragraph after image {i + 1}:");
        }

        // Save the original document.
        doc.Save(originalDoc);

        // -----------------------------------------------------------------
        // 3. Create a placeholder image that will replace all existing images.
        // -----------------------------------------------------------------
        CreateSampleImage(placeholderImage, 200, 200, Aspose.Drawing.Color.LightGray);

        // -----------------------------------------------------------------
        // 4. Load the document, replace each image with the placeholder.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalDoc);

        // Get all Shape nodes (including inline and floating images).
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Replace the image data with the placeholder image file.
                shape.ImageData.SetImage(placeholderImage);
            }
        }

        // Save the modified document.
        loadedDoc.Save(modifiedDoc);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(modifiedDoc))
            throw new InvalidOperationException($"Failed to create the modified document: {modifiedDoc}");

        // Clean up temporary files (optional).
        // File.Delete(sampleImage);
        // File.Delete(placeholderImage);
        // File.Delete(originalDoc);
    }

    // Helper method to create a deterministic PNG image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(backgroundColor);
            }

            // Save as PNG.
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
