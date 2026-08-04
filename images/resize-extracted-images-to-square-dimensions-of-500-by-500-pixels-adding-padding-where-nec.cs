using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string docPath = "sample.docx";

        // -------------------------------------------------
        // 1. Create a sample non‑square image (300x200)
        // -------------------------------------------------
        using (Bitmap bmp = new Bitmap(300, 200))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle to make the image recognizable
                g.FillRectangle(new SolidBrush(Color.Orange), 50, 50, 200, 100);
            }
            bmp.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Create a Word document and insert the image
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract images
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // -------------------------------------------------
            // 4. Get the image bytes from the shape
            // -------------------------------------------------
            byte[] imageBytes = shape.ImageData.ToByteArray();
            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure stream is at the beginning

                // -------------------------------------------------
                // 5. Load the image into Aspose.Drawing.Bitmap
                // -------------------------------------------------
                using (Bitmap original = new Bitmap(ms))
                {
                    // -------------------------------------------------
                    // 6. Create a 500x500 bitmap with white padding
                    // -------------------------------------------------
                    const int targetSize = 500;
                    using (Bitmap padded = new Bitmap(targetSize, targetSize))
                    {
                        using (Graphics g = Graphics.FromImage(padded))
                        {
                            g.Clear(Color.White);

                            // Compute scaling while preserving aspect ratio
                            double scale = Math.Min((double)targetSize / original.Width, (double)targetSize / original.Height);
                            int newWidth = (int)(original.Width * scale);
                            int newHeight = (int)(original.Height * scale);
                            int offsetX = (targetSize - newWidth) / 2;
                            int offsetY = (targetSize - newHeight) / 2;

                            // Draw the resized original image centered
                            g.DrawImage(original, offsetX, offsetY, newWidth, newHeight);
                        }

                        // -------------------------------------------------
                        // 7. Save the padded image
                        // -------------------------------------------------
                        string outputPath = $"resized_{imageIndex}.png";
                        padded.Save(outputPath);
                        Console.WriteLine($"Saved resized image: {outputPath}");
                    }
                }
            }

            imageIndex++;
        }

        // -------------------------------------------------
        // 8. Validation – ensure at least one image was saved
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Clean up temporary files (optional)
        // File.Delete(inputImagePath);
        // File.Delete(docPath);
    }
}
