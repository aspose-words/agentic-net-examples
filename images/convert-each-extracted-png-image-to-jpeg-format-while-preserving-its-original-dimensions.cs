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
        // Create a deterministic PNG image.
        const string pngPath = "sample.png";
        const int pngWidth = 200;
        const int pngHeight = 100;

        using (Bitmap bitmap = new Bitmap(pngWidth, pngHeight))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.LightBlue);
            bitmap.Save(pngPath, ImageFormat.Png);
        }

        // Create a Word document and insert the PNG image twice.
        const string docPath = "doc_with_png.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(pngPath);
        builder.InsertParagraph();
        builder.InsertImage(pngPath);
        doc.Save(docPath);

        // Load the document (already in memory) and extract PNG images.
        NodeCollection shapeNodes = doc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            if (shape.ImageData.ImageType != ImageType.Png)
                continue;

            // Get the PNG bytes.
            byte[] pngBytes = shape.ImageData.ToByteArray();

            // Load PNG bytes into Aspose.Drawing.Bitmap.
            using (MemoryStream ms = new MemoryStream(pngBytes))
            {
                ms.Position = 0; // Ensure stream is at the beginning.
                using (Bitmap pngBitmap = new Bitmap(ms))
                {
                    // Save as JPEG preserving original dimensions.
                    string jpegPath = $"extracted_{imageIndex}.jpg";
                    pngBitmap.Save(jpegPath, ImageFormat.Jpeg);

                    // Validate that the JPEG file was created.
                    if (!File.Exists(jpegPath))
                        throw new InvalidOperationException($"Failed to save JPEG image: {jpegPath}");

                    imageIndex++;
                }
            }
        }

        // Optional: indicate completion.
        Console.WriteLine($"Converted {imageIndex} PNG image(s) to JPEG format.");
    }
}
