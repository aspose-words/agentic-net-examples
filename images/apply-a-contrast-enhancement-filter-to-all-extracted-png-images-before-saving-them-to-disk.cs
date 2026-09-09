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
        string artifactsDir = "Artifacts";
        Directory.CreateDirectory(artifactsDir);

        // -----------------------------------------------------------------
        // 1. Create a sample PNG image that will be inserted into the document.
        // -----------------------------------------------------------------
        string sampleImagePath = Path.Combine(artifactsDir, "sample.png");
        using (Bitmap bitmap = new Bitmap(200, 200))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            // Fill background.
            graphics.Clear(Color.LightGray);
            // Draw a simple rectangle.
            graphics.DrawRectangle(new Pen(Color.Blue, 5), 20, 20, 160, 160);
            // Save the bitmap as PNG.
            bitmap.Save(sampleImagePath);
        }

        // -----------------------------------------------------------------
        // 2. Create a Word document and insert the PNG image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        string docPath = Path.Combine(artifactsDir, "DocumentWithImage.docx");
        doc.Save(docPath);

        // -----------------------------------------------------------------
        // 3. Load the document, find all PNG images, enhance contrast, and save them.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Png)
            {
                // Enhance contrast (value range 0.0 – 1.0, default 0.5).
                shape.ImageData.Contrast = 1.0; // maximum contrast

                // Save the modified image to disk.
                string outImagePath = Path.Combine(artifactsDir, $"extracted_{imageIndex}.png");
                shape.ImageData.Save(outImagePath);
                imageIndex++;
            }
        }

        // -----------------------------------------------------------------
        // 4. Validate that at least one image was saved.
        // -----------------------------------------------------------------
        if (imageIndex == 0)
            throw new InvalidOperationException("No PNG images were extracted from the document.");

        // The program finishes automatically.
    }
}
