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
        // Create a deterministic sample image to act as audio cover art.
        const string coverImagePath = "cover.png";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.White);
                // Draw a simple rectangle as placeholder content.
                g.DrawRectangle(new Pen(Aspose.Drawing.Color.Black), 20, 20, 160, 160);
            }
            bitmap.Save(coverImagePath);
        }

        // Create a new document and insert the sample image.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(coverImagePath);
        doc.Save(docPath);

        // Reload the document for extraction.
        Document loadedDoc = new Document(docPath);

        // Extract all images (including cover art) and save them as JPEG files.
        int extractedCount = 0;
        NodeCollection shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        foreach (Shape shape in shapes)
        {
            if (shape.HasImage)
            {
                byte[] imageBytes = shape.ImageData.ImageBytes;
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    using (Aspose.Drawing.Image img = Aspose.Drawing.Image.FromStream(ms))
                    {
                        string outputPath = $"extracted-{extractedCount}.jpg";
                        img.Save(outputPath, ImageFormat.Jpeg);
                        extractedCount++;
                    }
                }
            }
        }

        // Validate that at least one image was extracted.
        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");
    }
}
