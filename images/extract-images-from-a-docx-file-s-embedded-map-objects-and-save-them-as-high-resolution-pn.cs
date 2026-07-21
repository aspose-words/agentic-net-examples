using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample image (high‑resolution PNG)
        const int imgWidth = 800;
        const int imgHeight = 600;
        const string sampleImagePath = "sample.png";

        // Create bitmap and draw simple map‑like rectangle
        using (Bitmap bitmap = new Bitmap(imgWidth, imgHeight))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.FillRectangle(new SolidBrush(Color.LightBlue), 100, 100, 600, 400);
                g.DrawRectangle(new Pen(Color.DarkBlue, 5), 100, 100, 600, 400);
            }

            // Save the bitmap as PNG
            bitmap.Save(sampleImagePath);
        }

        // Create a DOCX document and insert the sample image as a shape
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // Reload the document to simulate extraction from an existing file
        Document loadedDoc = new Document(docPath);

        // Extract images from shape nodes (including map‑like objects) and save as high‑resolution PNG
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                string outFile = $"extracted-{extractedCount + 1}.png";
                shape.ImageData.Save(outFile);
                extractedCount++;
            }
        }

        // Validate that at least one image was extracted
        if (extractedCount == 0)
        {
            throw new InvalidOperationException("No images were extracted from the document.");
        }

        // Optional cleanup (commented out to keep output files for verification)
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
