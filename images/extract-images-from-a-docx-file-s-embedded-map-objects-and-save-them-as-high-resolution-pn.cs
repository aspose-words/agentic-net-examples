using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing; // Aspose.Drawing.Common provides Bitmap, Graphics, Color, Pen

public class ExtractMapImages
{
    public static void Main()
    {
        // Prepare output folder
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a sample high‑resolution PNG image that will act as a map.
        string mapImagePath = Path.Combine(outputDir, "map.png");
        CreateSampleMapImage(mapImagePath, 1200, 800); // high resolution

        // 2. Create a DOCX document and insert the map image.
        string docPath = Path.Combine(outputDir, "MapDocument.docx");
        CreateDocumentWithMap(docPath, mapImagePath);

        // 3. Load the document (no special load options needed for PNG).
        Document doc = new Document(docPath);

        // 4. Extract all images from shape nodes and save them as PNG files.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a file name for the extracted image.
            string extractedPath = Path.Combine(outputDir, $"extracted_image_{extractedCount}.png");

            // Save the image data. The original image is already PNG, so the saved file will be high‑resolution.
            shape.ImageData.Save(extractedPath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Optional: inform the user via console (no input required).
        Console.WriteLine($"Extracted {extractedCount} image(s) to \"{outputDir}\".");
    }

    // Creates a deterministic PNG file that represents a simple map.
    private static void CreateSampleMapImage(string filePath, int width, int height)
    {
        // Create a bitmap with the requested resolution.
        Bitmap bitmap = new Bitmap(width, height);
        Graphics graphics = Graphics.FromImage(bitmap);

        // Fill background.
        graphics.Clear(Color.White);

        // Draw a simple map‑like rectangle with a border.
        using (Pen pen = new Pen(Color.DarkBlue, 8))
        {
            graphics.DrawRectangle(pen, 50, 50, width - 100, height - 100);
        }

        // Save the bitmap as PNG.
        bitmap.Save(filePath);

        // Clean up resources.
        graphics.Dispose();
        bitmap.Dispose();
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithMap(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the map image into the document.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }
}
