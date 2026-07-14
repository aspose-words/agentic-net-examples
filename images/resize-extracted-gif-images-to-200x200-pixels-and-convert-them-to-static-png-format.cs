using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Define file paths
        const string sampleGifPath = "sample.gif";
        const string documentPath = "sample.docx";
        const string extractedGifPath = "extracted.gif";
        const string resizedPngPath = "resized.png";

        // -------------------------------------------------
        // 1. Create a sample GIF image (static, 300x300)
        // -------------------------------------------------
        using (Bitmap bitmap = new Bitmap(300, 300))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);
                // Optional: draw a simple rectangle for visual content
                graphics.DrawRectangle(Pens.Black, 50, 50, 200, 200);
            }
            bitmap.Save(sampleGifPath, ImageFormat.Gif);
        }

        // -------------------------------------------------
        // 2. Insert the GIF into a new Word document
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(sampleGifPath);
        doc.Save(documentPath);

        // -------------------------------------------------
        // 3. Load the document and extract the GIF image
        // -------------------------------------------------
        Document loadedDoc = new Document(documentPath);
        var shapes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        bool gifFound = false;

        foreach (Shape shape in shapes.OfType<Shape>())
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Gif)
            {
                shape.ImageData.Save(extractedGifPath);
                gifFound = true;
                break; // Assuming only one GIF for this example
            }
        }

        if (!gifFound)
            throw new Exception("No GIF image found in the document.");

        // -------------------------------------------------
        // 4. Resize the extracted GIF to 200x200 and convert to PNG
        // -------------------------------------------------
        using (Bitmap original = new Bitmap(extractedGifPath))
        {
            using (Bitmap resized = new Bitmap(200, 200))
            {
                using (Graphics graphics = Graphics.FromImage(resized))
                {
                    graphics.DrawImage(original, 0, 0, 200, 200);
                }
                resized.Save(resizedPngPath, ImageFormat.Png);
            }
        }

        // -------------------------------------------------
        // 5. Validate that the PNG was created
        // -------------------------------------------------
        if (!File.Exists(resizedPngPath))
            throw new Exception("Resized PNG image was not created.");

        // Cleanup temporary files (optional)
        // File.Delete(sampleGifPath);
        // File.Delete(documentPath);
        // File.Delete(extractedGifPath);
    }
}
