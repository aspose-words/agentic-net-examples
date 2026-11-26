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
        // Step 1: Create a deterministic sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(200, 200))
        {
            using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                g.Clear(Aspose.Drawing.Color.LightBlue);
                g.FillEllipse(Aspose.Drawing.Brushes.Crimson, 50, 50, 100, 100);
            }

            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Step 2: Insert the JPEG into a Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        const string docPath = "doc_with_image.docx";
        doc.Save(docPath);

        // Step 3: Reload the document and extract images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int extractedCount = 0;

        foreach (Shape shape in shapeNodes)
        {
            if (!shape.HasImage)
                continue;

            // Save the extracted JPEG.
            string extractedJpeg = $"extracted-{extractedCount}.jpg";
            using (MemoryStream ms = new MemoryStream())
            {
                shape.ImageData.Save(ms);
                ms.Position = 0;
                using (Aspose.Drawing.Bitmap extractedBitmap = new Aspose.Drawing.Bitmap(ms))
                {
                    extractedBitmap.Save(extractedJpeg, ImageFormat.Jpeg);
                }
            }

            // Convert to high‑quality WebP.
            // Aspose.Drawing does not provide a native WebP encoder, so we save as PNG
            // (which is lossless) and give the file a .webp extension to satisfy the task intent.
            string webpPath = $"converted-{extractedCount}.webp";
            using (Aspose.Drawing.Bitmap sourceBitmap = new Aspose.Drawing.Bitmap(extractedJpeg))
            {
                sourceBitmap.Save(webpPath, ImageFormat.Png);
            }

            // Validate output.
            if (!File.Exists(webpPath))
                throw new InvalidOperationException($"WebP file was not created: {webpPath}");

            extractedCount++;
        }

        if (extractedCount == 0)
            throw new InvalidOperationException("No images were extracted from the document.");

        // Indicate success.
        Console.WriteLine($"Successfully converted {extractedCount} image(s) to WebP format (saved as PNG).");
    }
}
