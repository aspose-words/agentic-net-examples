using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class BatchTiffToPngConverter
{
    public static void Main()
    {
        // Prepare input and output directories.
        string baseDir = Directory.GetCurrentDirectory();
        string inputDir = Path.Combine(baseDir, "InputImages");
        string outputDir = Path.Combine(baseDir, "OutputImages");
        Directory.CreateDirectory(inputDir);
        Directory.CreateDirectory(outputDir);

        // -----------------------------------------------------------------
        // Create deterministic sample TIFF images with distinct DPI values.
        // -----------------------------------------------------------------
        for (int i = 0; i < 3; i++)
        {
            string tiffPath = Path.Combine(inputDir, $"sample{i}.tiff");
            using (Bitmap bitmap = new Bitmap(200, 200))
            {
                // Set a distinct DPI for each image (72, 96, 120).
                float dpi = 72f + i * 24f;
                bitmap.SetResolution(dpi, dpi);

                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    g.Clear(Color.White);
                    using (Pen pen = new Pen(Color.Blue, 5))
                    {
                        g.DrawRectangle(pen, 20, 20, 160, 160);
                    }
                }

                // Save as TIFF (lossless).
                bitmap.Save(tiffPath, ImageFormat.Tiff);
            }
        }

        // --------------------------------------------------------------
        // Insert the created TIFF images into a Word document.
        // --------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        foreach (string tiffFile in Directory.GetFiles(inputDir, "*.tiff"))
        {
            builder.InsertImage(tiffFile);
            builder.Writeln(); // Separate images with a line break.
        }

        // Optional: save the document to demonstrate insertion.
        string docPath = Path.Combine(baseDir, "SampleDocument.docx");
        doc.Save(docPath, SaveFormat.Docx);

        // --------------------------------------------------------------
        // Extract each image from the document and convert it to PNG,
        // preserving the original DPI metadata.
        // --------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Save the image data to a memory stream.
            using (MemoryStream imageStream = new MemoryStream())
            {
                shape.ImageData.Save(imageStream);
                imageStream.Position = 0;

                // Load the image with Aspose.Drawing.Bitmap.
                using (Bitmap sourceBitmap = new Bitmap(imageStream))
                {
                    // Retrieve original DPI.
                    float originalDpiX = sourceBitmap.HorizontalResolution;
                    float originalDpiY = sourceBitmap.VerticalResolution;

                    // Create a new bitmap (clone) to ensure we have a writable instance.
                    using (Bitmap pngBitmap = new Bitmap(sourceBitmap))
                    {
                        // Preserve DPI metadata.
                        pngBitmap.SetResolution(originalDpiX, originalDpiY);

                        // Save as lossless PNG.
                        string pngPath = Path.Combine(outputDir, $"image{imageIndex}.png");
                        pngBitmap.Save(pngPath, ImageFormat.Png);
                    }
                }
            }

            imageIndex++;
        }

        // --------------------------------------------------------------
        // Validation: ensure at least one PNG was generated.
        // --------------------------------------------------------------
        int pngCount = Directory.GetFiles(outputDir, "*.png").Length;
        if (pngCount == 0)
            throw new InvalidOperationException("No PNG images were generated.");
    }
}
