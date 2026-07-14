using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Define file names.
        const string sampleImagePath = "sample.png";
        const string placeholderImagePath = "placeholder.png";
        const string originalDocPath = "original.docx";
        const string modifiedDocPath = "modified.docx";

        // -------------------------------------------------
        // Step 1: Create a sample image to be inserted.
        // -------------------------------------------------
        CreateSampleImage(sampleImagePath, 200, 100, Aspose.Drawing.Color.LightBlue, "Sample");

        // -------------------------------------------------
        // Step 2: Build a document that contains a few images.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert three sample images.
        builder.Writeln("First image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        builder.Writeln("Second image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        builder.Writeln("Third image:");
        builder.InsertImage(sampleImagePath);
        builder.Writeln();

        // Save the original document.
        doc.Save(originalDocPath);

        // -------------------------------------------------
        // Step 3: Create a placeholder image that will replace all existing images.
        // -------------------------------------------------
        CreateSampleImage(placeholderImagePath, 200, 100, Aspose.Drawing.Color.LightGray, "Placeholder");

        // -------------------------------------------------
        // Step 4: Load the document and replace each image with the placeholder.
        // -------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);

        // Get all Shape nodes (including images) in the document.
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>()
                              .Where(s => s.HasImage);

        foreach (Shape shape in shapes)
        {
            // Replace the image data with the placeholder image.
            shape.ImageData.SetImage(placeholderImagePath);
        }

        // Save the modified document.
        loadedDoc.Save(modifiedDocPath);

        // -------------------------------------------------
        // Step 5: Validate that the output file was created.
        // -------------------------------------------------
        if (!File.Exists(modifiedDocPath))
        {
            throw new InvalidOperationException($"Failed to create the modified document: {modifiedDocPath}");
        }

        // Clean up temporary image files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(placeholderImagePath);
    }

    // Helper method to create a deterministic bitmap image.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backColor, string text)
    {
        // Create a bitmap.
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            // Obtain a graphics object for drawing.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(backColor);

                // Optional: draw simple text in the center.
                // Note: Aspose.Drawing does not provide a direct DrawString method without a Font.
                // To keep the example simple and avoid font ambiguity, we skip drawing text.
            }

            // Save the bitmap to the specified file.
            bitmap.Save(filePath);
        }
    }
}
