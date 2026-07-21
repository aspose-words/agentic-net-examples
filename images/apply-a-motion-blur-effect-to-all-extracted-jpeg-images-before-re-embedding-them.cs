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
        // Directories for artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // 1. Create a deterministic sample JPEG image.
        string sampleImagePath = Path.Combine(artifactsDir, "sample.jpg");
        CreateSampleJpeg(sampleImagePath, 200, 200, Aspose.Drawing.Color.Red);

        // 2. Build a document that contains the sample JPEG image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string originalDocPath = Path.Combine(artifactsDir, "original.docx");
        doc.Save(originalDocPath);

        // 3. Load the document and process each JPEG image.
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            // Extract the original image bytes.
            byte[] originalBytes = shape.ImageData.ImageBytes;
            using (MemoryStream inputStream = new MemoryStream(originalBytes))
            using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromStream(inputStream))
            {
                // Apply a simple motion‑blur effect by drawing several shifted copies.
                using (Bitmap blurredBitmap = new Bitmap(originalImage.Width, originalImage.Height))
                using (Graphics graphics = Graphics.FromImage(blurredBitmap))
                {
                    // Clear background.
                    graphics.Clear(Aspose.Drawing.Color.Transparent);

                    // Number of steps and offset per step.
                    int steps = 5;
                    int offsetX = 2;

                    // Draw the image multiple times with incremental offset.
                    for (int i = 0; i < steps; i++)
                    {
                        int dx = i * offsetX;
                        graphics.DrawImage(originalImage, dx, 0, originalImage.Width, originalImage.Height);
                    }

                    // Save the blurred image to a stream.
                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        blurredBitmap.Save(outputStream, ImageFormat.Jpeg);
                        outputStream.Position = 0;

                        // Replace the shape's image with the blurred version.
                        shape.ImageData.SetImage(outputStream);
                    }
                }
            }
        }

        // 4. Save the modified document.
        string modifiedDocPath = Path.Combine(artifactsDir, "modified.docx");
        loadedDoc.Save(modifiedDocPath);

        // 5. Validate that the output file exists.
        if (!File.Exists(modifiedDocPath))
            throw new InvalidOperationException("The modified document was not saved correctly.");

        // Clean up temporary image file.
        if (File.Exists(sampleImagePath))
            File.Delete(sampleImagePath);
    }

    // Helper method to create a deterministic JPEG image.
    private static void CreateSampleJpeg(string filePath, int width, int height, Aspose.Drawing.Color fillColor)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(fillColor);
            bitmap.Save(filePath, ImageFormat.Jpeg);
        }
    }
}
