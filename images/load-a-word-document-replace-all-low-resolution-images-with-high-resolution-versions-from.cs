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
        // Create sample low‑resolution image (100x100)
        const string lowResPath = "low.png";
        using (Bitmap lowBmp = new Bitmap(100, 100))
        {
            using (Graphics g = Graphics.FromImage(lowBmp))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Red, 5))
                {
                    g.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            lowBmp.Save(lowResPath, ImageFormat.Png);
        }

        // Create sample high‑resolution image (500x500)
        const string highResPath = "high.png";
        using (Bitmap highBmp = new Bitmap(500, 500))
        {
            using (Graphics g = Graphics.FromImage(highBmp))
            {
                g.Clear(Color.White);
                using (Pen pen = new Pen(Color.Blue, 10))
                {
                    g.DrawEllipse(pen, 50, 50, 400, 400);
                }
            }
            highBmp.Save(highResPath, ImageFormat.Png);
        }

        // Create a Word document containing the low‑resolution image
        const string inputDocPath = "docLowRes.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(lowResPath);
        doc.Save(inputDocPath);

        // Load the document and replace low‑resolution images with high‑resolution ones
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int replacedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Load the image bytes into Aspose.Drawing.Image to inspect pixel dimensions
            using (MemoryStream imgStream = new MemoryStream(shape.ImageData.ImageBytes))
            {
                imgStream.Position = 0;
                using (Image img = Image.FromStream(imgStream))
                {
                    // Define low‑resolution threshold (e.g., width or height <= 150 pixels)
                    if (img.Width <= 150 && img.Height <= 150)
                    {
                        shape.ImageData.SetImage(highResPath);
                        replacedCount++;
                    }
                }
            }
        }

        const string outputDocPath = "docReplaced.docx";
        loadedDoc.Save(outputDocPath);

        // Validation
        if (!File.Exists(outputDocPath))
            throw new Exception($"Output document was not created: {outputDocPath}");

        if (replacedCount == 0)
            throw new Exception("No low‑resolution images were found to replace.");

        // Clean up temporary files (optional)
        // File.Delete(lowResPath);
        // File.Delete(highResPath);
        // File.Delete(inputDocPath);
    }
}
