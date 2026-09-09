using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class ReplaceLowResolutionImages
{
    public static void Main()
    {
        // Define file names.
        const string lowResImagePath = "lowres.png";
        const string highResImagePath = "highres.png";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create sample low‑resolution image (72 DPI, 100x100 pixels).
        // -----------------------------------------------------------------
        using (Bitmap lowResBitmap = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(lowResBitmap))
            {
                g.Clear(Color.White);
                // Draw a simple rectangle to make the image visible.
                g.DrawRectangle(Pens.Black, 10, 10, 80, 80);
            }

            // Set low DPI.
            lowResBitmap.SetResolution(72f, 72f);
            lowResBitmap.Save(lowResImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Create sample high‑resolution image (300 DPI, 500x500 pixels).
        // -----------------------------------------------------------------
        using (Bitmap highResBitmap = new Bitmap(500, 500))
        {
            using (Graphics g = Graphics.FromImage(highResBitmap))
            {
                g.Clear(Color.White);
                // Draw a larger rectangle.
                g.DrawRectangle(Pens.Blue, 50, 50, 400, 400);
            }

            // Set high DPI.
            highResBitmap.SetResolution(300f, 300f);
            highResBitmap.Save(highResImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 3. Create a Word document that contains the low‑resolution image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the low‑resolution image three times.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(lowResImagePath);
            builder.Writeln(); // Add a line break between images.
        }

        // Save the source document.
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document and replace low‑resolution images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);

        // Threshold DPI below which an image is considered low resolution.
        const double dpiThreshold = 150.0;

        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            ImageSize size = shape.ImageData.ImageSize;
            // If either horizontal or vertical DPI is below the threshold, replace the image.
            if (size.HorizontalResolution < dpiThreshold || size.VerticalResolution < dpiThreshold)
            {
                // Replace with the high‑resolution image.
                shape.ImageData.SetImage(highResImagePath);
            }
        }

        // Save the modified document.
        loadedDoc.Save(outputDocPath);

        // -----------------------------------------------------------------
        // 5. Validate that the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new InvalidOperationException($"Failed to create output document: {outputDocPath}");

        // Optional: clean up temporary files (comment out if you want to inspect them).
        // File.Delete(lowResImagePath);
        // File.Delete(highResImagePath);
        // File.Delete(inputDocPath);
    }
}
