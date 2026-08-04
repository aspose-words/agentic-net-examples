using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Words.Loading;
using Aspose.Drawing; // Aspose.Drawing.Common namespace

public class ExtractVideoFrameImages
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample high‑resolution image that will act as a video frame.
        string sampleImagePath = Path.Combine(artifactsDir, "frame.png");
        CreateSampleImage(sampleImagePath, 1920, 1080); // 1080p PNG

        // 2. Build a DOCX document and insert the sample image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the document and extract all images (including video frame images).
        Document doc = new Document(docPath);
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a PNG file name for the extracted image.
            string outFileName = $"extracted_{extractedCount}.png";
            string outPath = Path.Combine(artifactsDir, outFileName);

            // Save the image data. If the original format is not PNG, Aspose.Words will
            // convert it to PNG because we specify the .png extension.
            shape.ImageData.Save(outPath);
            extractedCount++;
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        Console.WriteLine($"Extracted {extractedCount} image(s) to folder: {artifactsDir}");
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                graphics.Clear(Color.White);
                // Optionally draw a simple rectangle to make the image visible.
                graphics.DrawRectangle(new Pen(Color.Black, 5), 10, 10, width - 20, height - 20);
            }

            // Save as PNG.
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image as an inline shape.
        builder.InsertImage(imagePath);

        // Save the document.
        doc.Save(docPath);
    }
}
