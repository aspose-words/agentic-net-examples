using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a deterministic PNG image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        const int imgWidth = 100;
        const int imgHeight = 100;

        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill background.
                graphics.Clear(Color.White);

                // Draw a simple red rectangle.
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    graphics.DrawRectangle(pen, 10, 10, imgWidth - 20, imgHeight - 20);
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Build a Word document that contains the PNG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the sample PNG twice.
        builder.InsertImage(sampleImagePath);
        builder.InsertParagraph();
        builder.InsertImage(sampleImagePath);

        // Save the document for reference (optional).
        string docPath = Path.Combine(artifactsDir, "sampleDoc.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Extract all PNG images, enhance contrast, and save them.
        // -----------------------------------------------------------------
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Apply a contrast enhancement (value between 0.0 and 1.0).
                shape.ImageData.Contrast = 0.8; // higher contrast

                // Determine file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(artifactsDir, $"extracted_{extractedCount}{extension}");

                // Save the modified image to disk.
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        // Validate that at least one PNG image was processed.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were found to extract.");

        // Program ends automatically; no user interaction required.
    }
}
