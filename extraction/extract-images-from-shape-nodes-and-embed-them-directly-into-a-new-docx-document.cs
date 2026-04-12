using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ExtractAndEmbedImages
{
    public static void Main()
    {
        // Create a folder for output files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Step 1: Create a sample source document with a few image shapes.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder srcBuilder = new DocumentBuilder(sourceDoc);

        // Insert first sample image (red square).
        using (MemoryStream imgStream = CreateSampleImage(100, 100, Color.Red))
        {
            srcBuilder.InsertImage(imgStream);
            srcBuilder.Writeln(); // separate images with a line break.
        }

        // Insert second sample image (green square).
        using (MemoryStream imgStream = CreateSampleImage(120, 80, Color.Green))
        {
            srcBuilder.InsertImage(imgStream);
            srcBuilder.Writeln();
        }

        // Save the source document (optional, for inspection).
        string sourcePath = Path.Combine(outputDir, "SourceDocument.docx");
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // Step 2: Extract images from shape nodes in the source document.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = sourceDoc.GetChildNodes(NodeType.Shape, true);

        // Prepare a new destination document where images will be embedded.
        Document destDoc = new Document();
        DocumentBuilder destBuilder = new DocumentBuilder(destDoc);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Get the raw image bytes from the shape.
                byte[] imageBytes = shape.ImageData.ToByteArray();

                // Insert the image into the destination document.
                destBuilder.InsertImage(imageBytes);
                destBuilder.Writeln(); // add a line break after each image.

                extractedCount++;
            }
        }

        // Validate that at least one image was extracted and embedded.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were found in the source document.");

        // -----------------------------------------------------------------
        // Step 3: Save the destination document containing the extracted images.
        // -----------------------------------------------------------------
        string destPath = Path.Combine(outputDir, "ExtractedImages.docx");
        destDoc.Save(destPath);

        // Simple confirmation (no interactive output required).
        Console.WriteLine($"Extraction complete. {extractedCount} image(s) embedded into '{destPath}'.");
    }

    // Helper method to create an in‑memory PNG image using Aspose.Drawing.
    private static MemoryStream CreateSampleImage(int width, int height, Color fillColor)
    {
        MemoryStream ms = new MemoryStream();
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(fillColor);
            }
            bitmap.Save(ms, ImageFormat.Png);
        }
        ms.Position = 0; // Reset stream position for reading.
        return ms;
    }
}
