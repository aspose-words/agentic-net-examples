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
        // Step 1: Create a sample PNG image (2000x1500) and save it locally.
        const string inputImagePath = "input.png";
        using (Bitmap bitmap = new Bitmap(2000, 1500))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to have some content.
                g.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, 1800, 1300);
            }
            bitmap.Save(inputImagePath, ImageFormat.Png);
        }

        // Step 2: Create a Word document and insert the PNG image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        doc.Save(docPath);

        // Step 3: Load the document and extract PNG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0;

                // Load the image into a bitmap.
                using (Bitmap originalBitmap = new Bitmap(originalStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    // Determine if resizing is needed.
                    int maxDimension = Math.Max(originalWidth, originalHeight);
                    if (maxDimension <= 1200)
                    {
                        // No resizing needed; just save the original image.
                        string outputPath = $"resized-{imageIndex}.png";
                        originalBitmap.Save(outputPath, ImageFormat.Png);
                        if (!File.Exists(outputPath))
                            throw new InvalidOperationException($"Failed to save image '{outputPath}'.");
                        imageIndex++;
                        continue;
                    }

                    // Calculate scaling factor.
                    double scale = 1200.0 / maxDimension;
                    int newWidth = (int)(originalWidth * scale);
                    int newHeight = (int)(originalHeight * scale);

                    // Create a new bitmap with the resized dimensions.
                    using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                    {
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.Clear(Color.Transparent);
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                        }

                        // Save the resized image.
                        string resizedPath = $"resized-{imageIndex}.png";
                        resizedBitmap.Save(resizedPath, ImageFormat.Png);
                        if (!File.Exists(resizedPath))
                            throw new InvalidOperationException($"Failed to save resized image '{resizedPath}'.");
                    }
                }
            }

            imageIndex++;
        }

        // Validation: ensure at least one resized image was produced.
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were found and processed in the document.");
    }
}
