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
        // Step 1: Create a sample non‑square image (200x300) and save it as input.png
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(200, 300))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple red rectangle for visual reference
                using (Pen pen = new Pen(Aspose.Drawing.Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, 180, 280);
                }
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // Step 2: Create a Word document and insert the sample image
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Step 3: Load the document (optional, we already have it) and extract images
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the original image to a memory stream
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // Reset before reading

                // Load the image into Aspose.Drawing.Bitmap
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    // Determine scaling factor to fit within 500x500 while preserving aspect ratio
                    const int targetSize = 500;
                    float widthScale = (float)targetSize / originalBitmap.Width;
                    float heightScale = (float)targetSize / originalBitmap.Height;
                    float scale = Math.Min(widthScale, heightScale);

                    int resizedWidth = (int)(originalBitmap.Width * scale);
                    int resizedHeight = (int)(originalBitmap.Height * scale);

                    // Create a new 500x500 bitmap with white background (padding)
                    using (Bitmap paddedBitmap = new Bitmap(targetSize, targetSize))
                    {
                        using (Graphics g = Graphics.FromImage(paddedBitmap))
                        {
                            g.Clear(Aspose.Drawing.Color.White);

                            // Calculate position to center the resized image
                            int offsetX = (targetSize - resizedWidth) / 2;
                            int offsetY = (targetSize - resizedHeight) / 2;

                            // Draw the resized image onto the padded bitmap
                            g.DrawImage(
                                originalBitmap,
                                new Rectangle(offsetX, offsetY, resizedWidth, resizedHeight));
                        }

                        // Save the padded image
                        string outputPath = $"resized-{imageIndex}.png";
                        paddedBitmap.Save(outputPath, ImageFormat.Png);
                        imageIndex++;
                    }
                }
            }
        }

        // Validation: ensure at least one resized image was created
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted and resized.");

        // Optional: clean up sample files (comment out if you want to inspect them)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
