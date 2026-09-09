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
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample PNG image to be used in the document.
        // -----------------------------------------------------------------
        const int sampleWidth = 100;
        const int sampleHeight = 100;
        const string sampleImagePath = "sample.png";

        using (Bitmap bmp = new Bitmap(sampleWidth, sampleHeight))
        using (Graphics gfx = Graphics.FromImage(bmp))
        {
            // Fill with white background.
            gfx.Clear(Aspose.Drawing.Color.White);
            // Draw a simple black rectangle inside the image.
            gfx.DrawRectangle(
                new Pen(Aspose.Drawing.Color.Black, 2),
                10, 10, sampleWidth - 20, sampleHeight - 20);

            // Save the sample PNG.
            bmp.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample PNG several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times to have multiple shapes to process.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // add a line break between images.
        }

        // -----------------------------------------------------------------
        // 3. Extract all PNG images, apply a 5‑pixel red border, and save them.
        // -----------------------------------------------------------------
        NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;
        const int borderSize = 5;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Obtain the raw image bytes.
            byte[] imageBytes = shape.ImageData.ToByteArray();

            // Load the image into an Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(imageBytes))
            using (Bitmap original = new Bitmap(ms))
            {
                // Create a new bitmap that is larger to accommodate the border.
                int newWidth = original.Width + borderSize * 2;
                int newHeight = original.Height + borderSize * 2;

                using (Bitmap bordered = new Bitmap(newWidth, newHeight))
                using (Graphics graphics = Graphics.FromImage(bordered))
                {
                    // Fill the whole bitmap with red – this becomes the border.
                    graphics.Clear(Aspose.Drawing.Color.Red);

                    // Draw the original image onto the new bitmap, offset by the border size.
                    graphics.DrawImage(
                        original,
                        borderSize,
                        borderSize,
                        original.Width,
                        original.Height);

                    // Save the resulting image.
                    string outFileName = $"Extracted_{extractedCount}.png";
                    bordered.Save(outFileName, ImageFormat.Png);
                }
            }

            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validation – ensure at least one image was written.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted and saved.");

        Console.WriteLine($"Successfully processed and saved {extractedCount} PNG image(s) with a red border.");
    }
}
