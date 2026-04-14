using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing; // Aspose.Drawing namespace for image creation

public class ReplaceLowResolutionImages
{
    public static void Main()
    {
        // File paths for sample images and documents
        const string lowResImagePath = "lowres.png";
        const string highResImagePath = "highres.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create sample low‑resolution image (72 DPI)
        // -----------------------------------------------------------------
        using (var lowBitmap = new Bitmap(100, 100))
        {
            using (var graphics = Graphics.FromImage(lowBitmap))
            {
                graphics.Clear(Color.White);
                // Draw a simple rectangle to make the image visible
                graphics.DrawRectangle(Pens.Black, 10, 10, 80, 80);
            }

            // Ensure the image file exists before using it
            lowBitmap.Save(lowResImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create sample high‑resolution image (300 DPI)
        // -----------------------------------------------------------------
        using (var highBitmap = new Bitmap(100, 100))
        {
            // Set a higher DPI to simulate a high‑resolution image
            // (SetResolution is available in Aspose.Drawing.Bitmap)
            highBitmap.SetResolution(300, 300);

            using (var graphics = Graphics.FromImage(highBitmap))
            {
                graphics.Clear(Color.White);
                graphics.DrawEllipse(Pens.Blue, 10, 10, 80, 80);
            }

            highBitmap.Save(highResImagePath);
        }

        // -----------------------------------------------------------------
        // 3. Build a document that contains low‑resolution images
        // -----------------------------------------------------------------
        var doc = new Document();
        var builder = new DocumentBuilder(doc);

        // Insert the low‑resolution image three times
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(lowResImagePath);
            builder.Writeln(); // Add a line break between images
        }

        // Save the source document
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document and replace low‑resolution images
        // -----------------------------------------------------------------
        var loadedDoc = new Document(inputDocPath);

        // Iterate over all Shape nodes in the document
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true)
                              .OfType<Shape>()
                              .Where(s => s.HasImage);

        foreach (var shape in shapes)
        {
            // Retrieve image size information
            ImageSize size = shape.ImageData.ImageSize;

            // Define a DPI threshold below which an image is considered low‑resolution
            const double dpiThreshold = 150.0;

            // If either horizontal or vertical resolution is below the threshold, replace it
            if (size.HorizontalResolution < dpiThreshold || size.VerticalResolution < dpiThreshold)
            {
                // Replace the image with the high‑resolution version
                shape.ImageData.SetImage(highResImagePath);
            }
        }

        // Save the modified document
        loadedDoc.Save(outputDocPath);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created successfully
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath) || new FileInfo(outputDocPath).Length == 0)
        {
            throw new InvalidOperationException("The output document was not created or is empty.");
        }

        // Clean up temporary image files (optional)
        // File.Delete(lowResImagePath);
        // File.Delete(highResImagePath);
    }
}
