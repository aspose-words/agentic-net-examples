using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Directories for artifacts
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic PNG image that will act as a video frame.
        string sampleImagePath = Path.Combine(artifactsDir, "frame.png");
        CreateSampleImage(sampleImagePath, 200, 200);

        // 2. Create a DOCX document and embed the sample image.
        string docPath = Path.Combine(artifactsDir, "sample.docx");
        CreateDocumentWithImage(docPath, sampleImagePath);

        // 3. Load the document and extract all images (including the embedded video frame).
        ExtractImagesFromDocument(docPath, artifactsDir);
    }

    // Creates a deterministic PNG image using Aspose.Drawing.
    private static void CreateSampleImage(string filePath, int width, int height)
    {
        // Ensure any previous file is removed.
        if (File.Exists(filePath))
            File.Delete(filePath);

        // Create bitmap and draw simple content.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            // Obtain a graphics object from the bitmap.
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background with white.
                g.Clear(Aspose.Drawing.Color.White);

                // Draw a blue rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
                {
                    g.DrawRectangle(pen, 10, 10, width - 20, height - 20);
                }

                // Draw some text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 24))
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black))
                {
                    g.DrawString("Frame", font, brush, new Aspose.Drawing.PointF(30, height / 2 - 15));
                }
            }

            // Save as PNG.
            bitmap.Save(filePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // Validate that the image was created.
        if (!File.Exists(filePath))
            throw new InvalidOperationException($"Failed to create sample image at '{filePath}'.");
    }

    // Creates a DOCX file and inserts the provided image.
    private static void CreateDocumentWithImage(string docPath, string imagePath)
    {
        // Ensure any previous file is removed.
        if (File.Exists(docPath))
            File.Delete(docPath);

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image (as a shape) into the document.
        Shape shape = builder.InsertImage(imagePath);
        shape.WrapType = WrapType.Inline;

        // Save the document.
        doc.Save(docPath, SaveFormat.Docx);

        // Validate that the document was created.
        if (!File.Exists(docPath))
            throw new InvalidOperationException($"Failed to create document at '{docPath}'.");
    }

    // Extracts all images from the document and saves them as high‑resolution PNG files.
    private static void ExtractImagesFromDocument(string docPath, string outputDir)
    {
        Document doc = new Document(docPath);

        // Get all shape nodes.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            // Process only shapes that actually contain image data.
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading.

                // Load the image into an Aspose.Drawing.Bitmap.
                using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imageStream))
                {
                    // Set a high resolution (e.g., 300 DPI).
                    bitmap.SetResolution(300, 300);

                    // Determine output file name.
                    string outFile = Path.Combine(outputDir, $"extracted_{imageIndex}.png");

                    // Save as PNG.
                    bitmap.Save(outFile, Aspose.Drawing.Imaging.ImageFormat.Png);

                    // Validate that the file was created.
                    if (!File.Exists(outFile))
                        throw new InvalidOperationException($"Failed to save extracted image to '{outFile}'.");

                    imageIndex++;
                }
            }
        }

        // Ensure at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
