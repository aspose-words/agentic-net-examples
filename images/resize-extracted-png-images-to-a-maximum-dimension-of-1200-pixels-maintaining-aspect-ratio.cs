using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        const string inputImagePath = "input.png";
        const string documentPath = "doc_with_image.docx";

        // -------------------------------------------------
        // 1. Create a sample PNG image larger than 1200px
        // -------------------------------------------------
        const int originalWidth = 2000;
        const int originalHeight = 1500;
        using (Bitmap bitmap = new Bitmap(originalWidth, originalHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Draw a simple rectangle to have some content
            graphics.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, originalWidth - 200, originalHeight - 200);
            bitmap.Save(inputImagePath);
        }

        // -------------------------------------------------
        // 2. Insert the image into a Word document
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract PNG images
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // -------------------------------------------------
            // 4. Obtain the image as a Bitmap
            // -------------------------------------------------
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0; // Reset before reading
                using (Bitmap originalBitmap = new Bitmap(ms))
                {
                    int width = originalBitmap.Width;
                    int height = originalBitmap.Height;

                    // -------------------------------------------------
                    // 5. Determine scaling factor to keep max dimension <= 1200
                    // -------------------------------------------------
                    const int maxDimension = 1200;
                    double scale = 1.0;
                    int maxCurrent = Math.Max(width, height);
                    if (maxCurrent > maxDimension)
                        scale = (double)maxDimension / maxCurrent;

                    int newWidth = (int)Math.Round(width * scale);
                    int newHeight = (int)Math.Round(height * scale);

                    // If no resizing needed, just save the original
                    if (scale >= 1.0)
                    {
                        string unchangedPath = $"extracted_{imageIndex}.png";
                        originalBitmap.Save(unchangedPath);
                        if (!File.Exists(unchangedPath))
                            throw new InvalidOperationException("Failed to save extracted image.");
                    }
                    else
                    {
                        // -------------------------------------------------
                        // 6. Resize the bitmap
                        // -------------------------------------------------
                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        using (Graphics g = Graphics.FromImage(resizedBitmap))
                        {
                            g.Clear(Color.Transparent);
                            g.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                            string resizedPath = $"resized_{imageIndex}.png";
                            resizedBitmap.Save(resizedPath);
                            if (!File.Exists(resizedPath))
                                throw new InvalidOperationException("Failed to save resized image.");
                        }
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one image was processed
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found to process.");
    }
}
