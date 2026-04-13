using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (300x200) and save it as input.png
        const int sampleWidth = 300;
        const int sampleHeight = 200;
        const string inputImagePath = "input.png";

        using (var bitmap = new Aspose.Drawing.Bitmap(sampleWidth, sampleHeight))
        {
            using (var graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.LightBlue);
                using (var pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, sampleWidth - 20, sampleHeight - 20);
                }
            }
            bitmap.Save(inputImagePath);
        }

        // Create a new document and insert the sample image
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(inputImagePath);
        const string docPath = "Original.docx";
        doc.Save(docPath);

        // Extract images from the document, resize them to 500x500 with padding, and save
        var shapes = doc.GetChildNodes(NodeType.Shape, true).OfType<Shape>();
        int extractedCount = 0;

        foreach (var shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Obtain raw image bytes from the shape
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap
            using (var sourceStream = new MemoryStream(imageBytes))
            {
                sourceStream.Position = 0; // Ensure stream is at the beginning
                using (var sourceBitmap = new Aspose.Drawing.Bitmap(sourceStream))
                {
                    const int targetSize = 500; // Desired square dimension
                    using (var squareBitmap = new Aspose.Drawing.Bitmap(targetSize, targetSize))
                    {
                        using (var graphics = Aspose.Drawing.Graphics.FromImage(squareBitmap))
                        {
                            // Fill background with white (padding)
                            graphics.Clear(Aspose.Drawing.Color.White);

                            // Compute scaling factor to fit the original image inside the square
                            double scale = Math.Min((double)targetSize / sourceBitmap.Width,
                                                    (double)targetSize / sourceBitmap.Height);
                            int newWidth = (int)(sourceBitmap.Width * scale);
                            int newHeight = (int)(sourceBitmap.Height * scale);

                            // Center the resized image within the square canvas
                            int offsetX = (targetSize - newWidth) / 2;
                            int offsetY = (targetSize - newHeight) / 2;

                            // Draw the resized image onto the canvas
                            graphics.DrawImage(sourceBitmap, offsetX, offsetY, newWidth, newHeight);
                        }

                        // Save the padded square image
                        string outputImagePath = $"resized_{extractedCount}.png";
                        squareBitmap.Save(outputImagePath);

                        // Validate that the file was created
                        if (!File.Exists(outputImagePath))
                            throw new InvalidOperationException($"Failed to save resized image: {outputImagePath}");
                    }
                }
            }

            extractedCount++;
        }

        // Ensure at least one image was processed
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
