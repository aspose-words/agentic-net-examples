using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create sample PNG image ----------
        string pngPath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                // Draw a simple red rectangle.
                g.FillRectangle(new SolidBrush(Color.Red), 20, 20, 160, 160);
            }
            bitmap.Save(pngPath);
        }

        // ---------- Create sample JPEG image (optional) ----------
        string jpgPath = Path.Combine(artifactsDir, "sample.jpg");
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillEllipse(new SolidBrush(Color.Blue), 20, 20, 160, 160);
            }
            bitmap.Save(jpgPath);
        }

        // ---------- Build a document containing the images ----------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert PNG image.
        builder.InsertImage(pngPath);
        builder.Writeln(); // separate

        // Insert JPEG image.
        builder.InsertImage(jpgPath);
        builder.Writeln();

        // Save the document.
        string docPath = Path.Combine(artifactsDir, "DocWithImages.docx");
        doc.Save(docPath);

        // ---------- Load the document and process PNG images ----------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        int extractedCount = 0;
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Process only PNG images.
            if (shape.ImageData.ImageType == ImageType.Png)
            {
                // Enhance contrast (range 0.0 – 1.0). 1.0 = maximum contrast.
                shape.ImageData.Contrast = 1.0;

                // Save the enhanced image.
                string outFile = Path.Combine(artifactsDir, $"Extracted_{imageIndex}.png");
                shape.ImageData.Save(outFile);
                extractedCount++;
                imageIndex++;
            }
        }

        // Validate that at least one PNG image was extracted and saved.
        if (extractedCount == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // Program ends automatically.
    }
}
