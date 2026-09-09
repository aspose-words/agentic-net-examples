using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;               // Aspose.Drawing namespace for graphics objects
using Aspose.Drawing.Imaging;      // For ImageFormat

public class Program
{
    public static void Main()
    {
        // Define deterministic file and folder names.
        const string workDir = "Work";
        const string mapImagePath = workDir + "/map.png";
        const string docPath = workDir + "/sample.docx";
        const string outputDir = workDir + "/Extracted";

        // Ensure required folders exist.
        Directory.CreateDirectory(workDir);
        Directory.CreateDirectory(outputDir);

        // -------------------------------------------------
        // 1. Create a sample high‑resolution PNG image.
        // -------------------------------------------------
        const int imgWidth = 1200;
        const int imgHeight = 800;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill background.
                g.Clear(Aspose.Drawing.Color.LightBlue);

                // Draw a simple map‑like rectangle.
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.DarkBlue, 5))
                {
                    g.DrawRectangle(pen, 100, 100, imgWidth - 200, imgHeight - 200);
                }

                // Add some text.
                using (Aspose.Drawing.Font font = new Aspose.Drawing.Font("Arial", 48, Aspose.Drawing.FontStyle.Bold))
                using (Aspose.Drawing.Brush brush = new Aspose.Drawing.SolidBrush(Aspose.Drawing.Color.DarkRed))
                {
                    g.DrawString("Sample Map", font, brush, new Aspose.Drawing.PointF(250, 350));
                }
            }

            // Save the bitmap as PNG.
            bitmap.Save(mapImagePath, Aspose.Drawing.Imaging.ImageFormat.Png);
        }

        // -------------------------------------------------
        // 2. Create a DOCX document and embed the PNG image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(mapImagePath); // Insert the image as an inline shape.
        doc.Save(docPath);

        // -------------------------------------------------
        // 3. Load the document and extract all embedded images.
        // -------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Force PNG extension for the extracted file.
            string extractedPath = Path.Combine(outputDir, $"ExtractedImage_{imageIndex}.png");

            // Save the image data directly to the file.
            shape.ImageData.Save(extractedPath);

            // Validate that the file was created.
            if (!File.Exists(extractedPath))
                throw new InvalidOperationException($"Failed to save image {extractedPath}");

            Console.WriteLine($"Extracted image saved to: {extractedPath}");
            imageIndex++;
        }

        // Ensure at least one image was extracted.
        if (imageIndex == 0)
            throw new InvalidOperationException("No images were found in the document.");

        Console.WriteLine("Image extraction completed successfully.");
    }
}
