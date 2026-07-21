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
        // Step 1: Create a deterministic sample image (300x200) and save it as input.png
        const int sampleWidth = 300;
        const int sampleHeight = 200;
        string inputImagePath = "input.png";

        using (Bitmap sampleBitmap = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics g = Graphics.FromImage(sampleBitmap))
        {
            g.Clear(Color.LightBlue);
            // Draw a simple rectangle to make the image recognizable
            g.DrawRectangle(Pens.Black, 10, 10, sampleWidth - 20, sampleHeight - 20);
            sampleBitmap.Save(inputImagePath);
        }

        // Step 2: Create a Word document and insert the sample image
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        string docPath = "sample.docx";
        doc.Save(docPath);

        // Step 3: Load the document (optional, we already have it) and extract images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Obtain image bytes from the shape
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the original image into Aspose.Drawing.Bitmap via a MemoryStream
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap original = new Bitmap(ms))
            {
                // Desired square size
                const int targetSize = 500;

                // Calculate scaling factor to fit the original image within the square while preserving aspect ratio
                double scale = Math.Min((double)targetSize / original.Width, (double)targetSize / original.Height);
                int scaledWidth = (int)(original.Width * scale);
                int scaledHeight = (int)(original.Height * scale);

                // Create a new bitmap with white background
                using (Bitmap squareBitmap = new Bitmap(targetSize, targetSize))
                using (Graphics graphics = Graphics.FromImage(squareBitmap))
                {
                    graphics.Clear(Color.White);

                    // Center the scaled image within the square
                    int offsetX = (targetSize - scaledWidth) / 2;
                    int offsetY = (targetSize - scaledHeight) / 2;

                    // Draw the scaled image onto the square canvas
                    graphics.DrawImage(
                        original,
                        new Rectangle(offsetX, offsetY, scaledWidth, scaledHeight));

                    // Save the resized image
                    string outputPath = $"extracted_{extractedCount}.png";
                    squareBitmap.Save(outputPath, ImageFormat.Png);

                    // Validate that the file was created
                    if (!File.Exists(outputPath))
                        throw new InvalidOperationException($"Failed to create output image: {outputPath}");

                    extractedCount++;
                }
            }
        }

        // Validation: ensure at least one image was processed
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Cleanup: optional removal of temporary files (commented out to keep results)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
