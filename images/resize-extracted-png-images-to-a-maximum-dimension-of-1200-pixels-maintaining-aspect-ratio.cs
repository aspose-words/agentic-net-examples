using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample PNG image larger than the target size.
        const string sampleImagePath = "sample.png";
        CreateSamplePng(sampleImagePath, 2000, 1500); // 2000x1500 pixels

        // Build a simple document that contains the sample image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "DocumentWithImage.docx";
        doc.Save(docPath);

        // Extract all PNG images from the document.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage || shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0; // reset before reading

                // Load the image into a bitmap for processing.
                using (Bitmap originalBitmap = new Bitmap(imageStream))
                {
                    int originalWidth = originalBitmap.Width;
                    int originalHeight = originalBitmap.Height;

                    const double maxDimension = 1200.0;
                    double scale = 1.0;

                    // Determine scaling factor if either dimension exceeds the maximum.
                    if (originalWidth > maxDimension || originalHeight > maxDimension)
                    {
                        scale = maxDimension / Math.Max(originalWidth, originalHeight);
                    }

                    string outputFileName = $"Resized_{extractedCount}.png";

                    if (scale < 1.0)
                    {
                        // Calculate new dimensions while preserving aspect ratio.
                        int newWidth = (int)(originalWidth * scale);
                        int newHeight = (int)(originalHeight * scale);

                        // Create a new bitmap with the target size and draw the scaled image.
                        using (Bitmap resizedBitmap = new Bitmap(newWidth, newHeight))
                        using (Graphics graphics = Graphics.FromImage(resizedBitmap))
                        {
                            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            graphics.DrawImage(originalBitmap, 0, 0, newWidth, newHeight);
                            resizedBitmap.Save(outputFileName, ImageFormat.Png);
                        }
                    }
                    else
                    {
                        // Image already within limits; save it unchanged.
                        originalBitmap.Save(outputFileName, ImageFormat.Png);
                    }

                    extractedCount++;
                }
            }
        }

        // Validation: ensure at least one PNG image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");
    }

    // Helper: creates a deterministic PNG file with simple graphics.
    private static void CreateSamplePng(string filePath, int width, int height)
    {
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            using (Pen pen = new Pen(Color.Red, 5))
            {
                graphics.DrawRectangle(pen, 10, 10, width - 20, height - 20);
            }
            bitmap.Save(filePath, ImageFormat.Png);
        }
    }
}
