using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;
using Aspose.Drawing.Drawing2D;    // For InterpolationMode

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic sample JPEG image (1200x900).
        // -----------------------------------------------------------------
        const string sampleImagePath = "sample.jpg";
        const int sampleWidth = 1200;
        const int sampleHeight = 900;

        // Create bitmap and fill with white background.
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(sampleWidth, sampleHeight))
        using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
        {
            graphics.Clear(Aspose.Drawing.Color.White);
            // Draw simple text.
            graphics.DrawString(
                "Sample JPEG",
                new Aspose.Drawing.Font("Arial", 48),
                Aspose.Drawing.Brushes.Black,
                new Aspose.Drawing.PointF(100, 400));

            // Save as JPEG.
            bitmap.Save(sampleImagePath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document and insert the sample JPEG several times.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image three times to have multiple JPEG shapes.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertImage(sampleImagePath);
            builder.Writeln(); // add a line break between images.
        }

        const string originalDocPath = "Original.docx";
        doc.Save(originalDocPath);

        // -----------------------------------------------------------------
        // 3. Load the document and process all JPEG images.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(originalDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int jpegCount = 0;
        int resizedCount = 0;

        // Ensure output directory exists.
        const string outputDir = "Output";
        Directory.CreateDirectory(outputDir);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only JPEG images.
            if (shape.ImageData.ImageType != ImageType.Jpeg)
                continue;

            jpegCount++;

            // Save the original image to a memory stream.
            using (MemoryStream originalStream = new MemoryStream())
            {
                shape.ImageData.Save(originalStream);
                originalStream.Position = 0; // reset before reading.

                // Load the image using Aspose.Drawing.
                using (Aspose.Drawing.Image originalImage = Aspose.Drawing.Image.FromStream(originalStream))
                {
                    // If width already <= 800, just save the original.
                    if (originalImage.Width <= 800)
                    {
                        string outPath = Path.Combine(outputDir, $"Image_{jpegCount}_original.jpg");
                        originalImage.Save(outPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        continue;
                    }

                    // Calculate new dimensions preserving aspect ratio.
                    int newWidth = 800;
                    int newHeight = (int)(originalImage.Height * (800.0 / originalImage.Width));

                    // Create a new bitmap with the target size.
                    using (Aspose.Drawing.Bitmap resizedBitmap = new Aspose.Drawing.Bitmap(newWidth, newHeight))
                    using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(resizedBitmap))
                    {
                        // High quality resizing.
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(originalImage, 0, 0, newWidth, newHeight);

                        // Save the resized image.
                        string resizedPath = Path.Combine(outputDir, $"Image_{jpegCount}_resized.jpg");
                        resizedBitmap.Save(resizedPath, Aspose.Drawing.Imaging.ImageFormat.Jpeg);
                        resizedCount++;
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 4. Validation.
        // -----------------------------------------------------------------
        if (jpegCount == 0)
            throw new InvalidOperationException("No JPEG images were found in the document.");

        Console.WriteLine($"Total JPEG images found: {jpegCount}");
        Console.WriteLine($"Images resized and saved: {resizedCount}");
        Console.WriteLine($"All output files are located in the \"{outputDir}\" folder.");
    }
}
