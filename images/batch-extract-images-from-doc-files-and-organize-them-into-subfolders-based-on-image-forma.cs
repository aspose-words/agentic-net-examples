using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create deterministic sample images (PNG, JPEG, BMP).
        CreateSampleImage("sample1.png", 200, 200, Aspose.Drawing.Color.LightBlue);
        CreateSampleImage("sample2.jpg", 200, 200, Aspose.Drawing.Color.LightCoral);
        CreateSampleImage("sample3.bmp", 200, 200, Aspose.Drawing.Color.LightGreen);

        // Step 2: Build a DOCX document and insert the sample images.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        builder.Writeln("Document with sample images:");
        builder.InsertImage("sample1.png");
        builder.InsertParagraph();
        builder.InsertImage("sample2.jpg");
        builder.InsertParagraph();
        builder.InsertImage("sample3.bmp");

        // Save the document to the local file system.
        const string docPath = "SampleDocument.docx";
        doc.Save(docPath);

        // Step 3: Load the document (reuse the same instance is also fine).
        Document loadedDoc = new Document(docPath);

        // Step 4: Extract all images from the document.
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine the image type and corresponding file extension.
            string extension = Aspose.Words.FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
            string formatFolder = extension.TrimStart('.').ToLowerInvariant();

            // Ensure the subfolder exists.
            Directory.CreateDirectory(formatFolder);

            // Build a deterministic file name.
            string imageFileName = $"Image_{imageIndex}{extension}";
            string outputPath = Path.Combine(formatFolder, imageFileName);

            // Save the image data to the file system.
            shape.ImageData.Save(outputPath);
            extractedCount++;
            imageIndex++;
        }

        // Validation: at least one image must have been extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Summary output (non‑interactive).
        Console.WriteLine($"Extracted {extractedCount} image(s) into format‑based subfolders.");
    }

    // Helper method to create a deterministic bitmap using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height, Aspose.Drawing.Color backgroundColor)
    {
        // Create a bitmap with the specified dimensions.
        Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height);
        // Obtain a graphics object for drawing.
        Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap);
        // Fill the bitmap with a solid background color.
        graphics.Clear(backgroundColor);
        // Save the bitmap to the specified file.
        bitmap.Save(filePath);
        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }
}
