using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Prepare folders.
        string baseDir = Directory.GetCurrentDirectory();
        string artifactsDir = Path.Combine(baseDir, "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample image using Aspose.Drawing.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle to make the image non‑empty.
                g.FillRectangle(new SolidBrush(Aspose.Drawing.Color.Blue), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath, ImageFormat.Png);
        }

        // -----------------------------------------------------------------
        // 2. Build a DOCM document that contains the image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the image; the returned Shape represents the picture.
        Shape pictureShape = builder.InsertImage(sampleImagePath);
        pictureShape.Name = "MyEmbeddedImage"; // Give the shape a deterministic name.

        // Save the document as a macro‑enabled file.
        string docmPath = Path.Combine(artifactsDir, "sample.docm");
        doc.Save(docmPath, SaveFormat.Docm);

        // -----------------------------------------------------------------
        // 3. Load the DOCM file and extract each embedded image.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docmPath);

        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine a file extension that matches the image type.
            string extension = FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType);

            // Use the shape's name if present; otherwise fall back to an index‑based name.
            string baseName = !string.IsNullOrEmpty(shape.Name) ? shape.Name : $"Image_{extractedCount}";
            string outputPath = Path.Combine(artifactsDir, $"{baseName}{extension}");

            // Save the image data to the file system.
            shape.ImageData.Save(outputPath);
            extractedCount++;
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one image was extracted.
        // -----------------------------------------------------------------
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // The program finishes automatically; no user interaction is required.
    }
}
