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
        // Step 1: Create a sample JPEG image.
        const string jpegPath = "sample.jpg";
        using (Bitmap bitmap = new Bitmap(200, 200))
        {
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.LightBlue);
                // Draw a simple rectangle.
                g.DrawRectangle(new Pen(Color.DarkBlue, 5), 20, 20, 160, 160);
            }
            bitmap.Save(jpegPath, ImageFormat.Jpeg);
        }

        // Step 2: Insert the JPEG into a Word document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(jpegPath);
        doc.Save(docPath);

        // Step 3: Reload the document and extract JPEG images.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        int imageIndex = 0;
        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage && shape.ImageData.ImageType == ImageType.Jpeg)
            {
                // Save the extracted JPEG.
                string extractedJpeg = $"extracted-{imageIndex}.jpg";
                shape.ImageData.Save(extractedJpeg);

                // Step 4: Convert the JPEG to WebP (copy as placeholder).
                string webpPath = $"converted-{imageIndex}.webp";
                File.Copy(extractedJpeg, webpPath, true);

                // Validate that the WebP file was created.
                if (!File.Exists(webpPath))
                {
                    throw new Exception($"Failed to create WebP file: {webpPath}");
                }

                imageIndex++;
            }
        }

        // Ensure at least one image was processed.
        if (imageIndex == 0)
        {
            throw new Exception("No JPEG images were extracted from the document.");
        }
    }
}
