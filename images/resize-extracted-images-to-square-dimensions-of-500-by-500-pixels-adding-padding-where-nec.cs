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
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string documentPath = "sample.docx";

        // -------------------------------------------------
        // 1. Create a sample image (300x200) using Aspose.Drawing
        // -------------------------------------------------
        using (Bitmap sampleBitmap = new Bitmap(300, 200))
        {
            using (Graphics g = Graphics.FromImage(sampleBitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple ellipse to make the image recognizable
                g.FillEllipse(Brushes.Orange, 50, 30, 200, 140);
            }
            sampleBitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the sample image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract each image,
        //    resize it to 500x500 with padding, and save.
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .Cast<Shape>()
                              .Where(s => s.HasImage)
                              .ToList();

        if (!shapes.Any())
            throw new InvalidOperationException("No images were found in the document.");

        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            // Get original image bytes
            byte[] originalBytes = shape.ImageData.ToByteArray();

            // Load original image into a Bitmap
            using (MemoryStream originalStream = new MemoryStream(originalBytes))
            using (Bitmap originalBitmap = new Bitmap(originalStream))
            {
                // Desired square size
                const int targetSize = 500;

                // Compute scaling factor to fit within the square while preserving aspect ratio
                double scale = Math.Min((double)targetSize / originalBitmap.Width,
                                        (double)targetSize / originalBitmap.Height);

                int scaledWidth = (int)(originalBitmap.Width * scale);
                int scaledHeight = (int)(originalBitmap.Height * scale);

                // Offsets to center the image
                int offsetX = (targetSize - scaledWidth) / 2;
                int offsetY = (targetSize - scaledHeight) / 2;

                // Create a new square bitmap with white background
                using (Bitmap squareBitmap = new Bitmap(targetSize, targetSize))
                {
                    using (Graphics g = Graphics.FromImage(squareBitmap))
                    {
                        g.Clear(Color.White);
                        // Draw the scaled original image onto the square canvas
                        g.DrawImage(originalBitmap, offsetX, offsetY, scaledWidth, scaledHeight);
                    }

                    // Save the resized image
                    string resizedImagePath = $"resized_{imageIndex}.png";
                    squareBitmap.Save(resizedImagePath, ImageFormat.Png);

                    // Validate that the file was created
                    if (!File.Exists(resizedImagePath))
                        throw new InvalidOperationException($"Failed to create {resizedImagePath}.");

                    // Optional: replace the image in the document with the resized version
                    // shape.ImageData.SetImage(squareBitmap);
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 4. (Optional) Save the document after replacement
        // -------------------------------------------------
        // loadedDoc.Save("sample_resized.docx");
    }
}
