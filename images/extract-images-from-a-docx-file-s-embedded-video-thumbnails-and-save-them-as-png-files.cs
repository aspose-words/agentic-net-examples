using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a deterministic sample thumbnail image.
        const string sampleImagePath = "thumb.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                // Draw a simple rectangle.
                g.DrawRectangle(new Pen(Aspose.Drawing.Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(sampleImagePath);
        }

        // Create a DOCX document and insert the sample image (simulating a video thumbnail).
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleImagePath);
        doc.Save(docPath);

        // Load the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Iterate through all Shape nodes and extract images.
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapes)
        {
            if (!shape.HasImage)
                continue;

            // Obtain the raw image bytes.
            byte[] imageBytes = shape.ImageData.ImageBytes;

            using (MemoryStream ms = new MemoryStream(imageBytes))
            {
                ms.Position = 0; // Ensure the stream is at the beginning.
                using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(ms))
                {
                    string outputPath = $"extracted-{extractedCount + 1}.png";
                    img.Save(outputPath, ImageFormat.Png);
                    extractedCount++;
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Cleanup sample files (optional).
        // File.Delete(sampleImagePath);
        // File.Delete(docPath);
    }
}
