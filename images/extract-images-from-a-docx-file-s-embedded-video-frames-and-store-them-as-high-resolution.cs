using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class ExtractVideoFrameImages
{
    public static void Main()
    {
        // Prepare folders
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a sample high‑resolution PNG image that will act as a video frame.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        CreateSampleImage(sampleImagePath, 300, 300);

        // 2. Build a DOCX document and insert the sample image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the document and extract all images (including the inserted video‑frame image).
        ExtractImagesFromDocument(docPath, artifactsDir);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // White background
            graphics.Clear(Aspose.Drawing.Color.White);

            // Draw a red ellipse to make the image recognizable
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 5))
            {
                graphics.DrawEllipse(pen, 10, 10, width - 20, height - 20);
            }

            // Save as PNG
            bitmap.Save(filePath);
        }
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image (simulating an embedded video frame)
        builder.InsertImage(imagePath);

        // Save the document
        doc.Save(docPath);
    }

    // Extracts all images from the document and saves them as high‑resolution PNG files.
    private static void ExtractImagesFromDocument(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Retrieve the raw image bytes from the shape.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the bytes into a bitmap so we can control the output format and resolution.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap bitmap = new Bitmap(ms))
            {
                // Set a high resolution (e.g., 300 DPI) for the exported PNG.
                bitmap.SetResolution(300, 300);

                string outFile = Path.Combine(outputDir, $"extracted_{imageIndex}.png");
                bitmap.Save(outFile);
                imageIndex++;
            }
        }

        // Validation: ensure at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
