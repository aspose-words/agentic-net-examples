using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing; // Provides Bitmap, Graphics, Color, etc.

public class ExtractMapImages
{
    public static void Main()
    {
        // Define deterministic file names and folders.
        string workDir = Path.Combine(Directory.GetCurrentDirectory(), "Work");
        Directory.CreateDirectory(workDir);
        string mapImagePath = Path.Combine(workDir, "map.png");
        string docPath = Path.Combine(workDir, "sample.docx");
        string outputDir = Path.Combine(workDir, "Extracted");
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a sample high‑resolution PNG image.
        // -------------------------------------------------
        int width = 1200;   // high resolution width
        int height = 800;   // high resolution height
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(width, height))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple map‑like rectangle.
                g.FillRectangle(new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.LightBlue), 100, 100, width - 200, height - 200);
                g.DrawRectangle(new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5), 100, 100, width - 200, height - 200);
                // Add some text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48))
                {
                    g.DrawString("Sample Map", font, new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.Black),
                        new Aspose.Drawing.PointF(width / 2 - 150, height / 2 - 24));
                }
            }
            // Save the bitmap as PNG.
            bitmap.Save(mapImagePath);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and embed the PNG image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert the image as an inline shape.
        builder.InsertImage(mapImagePath);
        // Save the document.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all embedded images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage)
            {
                // Determine appropriate file extension based on image type.
                string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);
                string outFile = Path.Combine(outputDir, $"extracted_{imageIndex}{extension}");
                // Save the image data.
                shape.ImageData.Save(outFile);
                imageIndex++;
            }
        }

        // -------------------------------------------------
        // 4. Validation – ensure at least one image was extracted.
        // -------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
