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
        // Directories for temporary files.
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);
        string lowResImagePath = Path.Combine(artifactsDir, "low.png");
        string highResImagePath = Path.Combine(artifactsDir, "high.png");
        string inputDocPath = Path.Combine(artifactsDir, "input.docx");
        string outputDocPath = Path.Combine(artifactsDir, "output.docx");

        // -------------------------------------------------
        // 1. Create sample low‑resolution image (100x100).
        // -------------------------------------------------
        using (Bitmap lowBitmap = new Bitmap(100, 100))
        using (Graphics g = Graphics.FromImage(lowBitmap))
        {
            g.Clear(Color.White);
            // Draw a simple rectangle to make the image visible.
            g.DrawRectangle(new Pen(Color.Black, 2), 10, 10, 80, 80);
            lowBitmap.Save(lowResImagePath);
        }

        // -------------------------------------------------
        // 2. Create sample high‑resolution image (500x500).
        // -------------------------------------------------
        using (Bitmap highBitmap = new Bitmap(500, 500))
        using (Graphics g = Graphics.FromImage(highBitmap))
        {
            g.Clear(Color.White);
            g.DrawRectangle(new Pen(Color.Blue, 5), 20, 20, 460, 460);
            highBitmap.Save(highResImagePath);
        }

        // -------------------------------------------------
        // 3. Build a document that contains low‑resolution images.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        // Insert three low‑resolution images.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertParagraph();
            builder.InsertImage(lowResImagePath);
        }
        doc.Save(inputDocPath);

        // -------------------------------------------------
        // 4. Load the document and replace low‑resolution images.
        // -------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Retrieve image size in pixels.
            ImageSize imgSize = shape.ImageData.ImageSize;
            // Define a threshold for low resolution (e.g., width or height < 200 px).
            if (imgSize.WidthPixels < 200 || imgSize.HeightPixels < 200)
            {
                // Replace with the high‑resolution image.
                shape.ImageData.SetImage(highResImagePath);
                replacedCount++;
            }
        }

        // Validate that at least one image was replaced.
        if (replacedCount == 0)
            throw new InvalidOperationException("No low‑resolution images were found to replace.");

        // Save the modified document.
        loadedDoc.Save(outputDocPath);

        // -------------------------------------------------
        // 5. Final validation.
        // -------------------------------------------------
        if (!File.Exists(outputDocPath))
            throw new FileNotFoundException("The output document was not created.", outputDocPath);

        // Optionally, report the result (no console interaction required).
        // The program ends here.
    }
}
