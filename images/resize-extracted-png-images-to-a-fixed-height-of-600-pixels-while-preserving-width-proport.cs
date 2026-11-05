using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Step 1: Create a deterministic sample PNG image.
        const string sampleImagePath = "input.png";
        using (var bitmap = new Bitmap(200, 200))
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                graphics.FillRectangle(new SolidBrush(Color.Red), 50, 50, 100, 100);
            }
            bitmap.Save(sampleImagePath);
        }

        // Step 2: Create a Word document and insert the sample image.
        var doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Step 3: Reload the document for extraction.
        var loadedDoc = new Document(docPath);

        // Step 4: Extract PNG images, resize them to height 600px while preserving aspect ratio.
        var shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        for (int i = 0; i < shapeNodes.Count; i++)
        {
            var shape = (Shape)shapeNodes[i];
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Save the extracted image.
            string extractedPath = $"extracted_{i}.png";
            shape.ImageData.Save(extractedPath);
            extractedCount++;

            // Load the extracted image for resizing.
            using (var originalBitmap = new Bitmap(extractedPath))
            {
                int originalWidth = originalBitmap.Width;
                int originalHeight = originalBitmap.Height;

                // Desired fixed height.
                const int targetHeight = 600;
                // Calculate proportional width.
                int targetWidth = (int)Math.Round(originalWidth * (targetHeight / (double)originalHeight));

                // Create a new bitmap with the target dimensions.
                using (var resizedBitmap = new Bitmap(targetWidth, targetHeight))
                {
                    using (var graphics = Graphics.FromImage(resizedBitmap))
                    {
                        graphics.Clear(Color.Transparent);
                        graphics.DrawImage(originalBitmap, 0, 0, targetWidth, targetHeight);
                    }

                    // Save the resized image.
                    string resizedPath = $"resized_{i}.png";
                    resizedBitmap.Save(resizedPath);

                    // Validate that the resized file exists.
                    if (!File.Exists(resizedPath))
                        throw new InvalidOperationException($"Resized image was not saved: {resizedPath}");
                }
            }

            // Validate that the extracted file exists.
            if (!File.Exists(extractedPath))
                throw new InvalidOperationException($"Extracted image was not saved: {extractedPath}");
        }

        // Validation: ensure at least one PNG image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were found and processed in the document.");

        // Cleanup: optional removal of intermediate files can be added here.
    }
}
