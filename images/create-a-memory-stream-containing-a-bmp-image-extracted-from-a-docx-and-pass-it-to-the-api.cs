using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // 1. Create a deterministic BMP image file.
        const string bmpPath = "input.bmp";
        const int imgWidth = 100;
        const int imgHeight = 100;

        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                graphics.Clear(Aspose.Drawing.Color.White);
                using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Black))
                {
                    graphics.DrawRectangle(pen, 10, 10, 80, 80);
                }
            }
            bitmap.Save(bmpPath);
        }

        // 2. Insert the BMP into a new DOCX document.
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        doc.Save(docPath);

        // 3. Load the document and locate the first shape that contains an image.
        Document loadedDoc = new Document(docPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);
        Shape imageShape = null;

        foreach (Shape shape in shapeNodes)
        {
            if (shape.HasImage)
            {
                imageShape = shape;
                break;
            }
        }

        if (imageShape == null)
            throw new InvalidOperationException("No image found in the document.");

        // 4. Save the image to a memory stream (BMP format is preserved).
        using (MemoryStream imageStream = new MemoryStream())
        {
            imageShape.ImageData.Save(imageStream);
            imageStream.Position = 0; // Reset before any further use.

            // 5. Example API usage: write the stream to a file to verify extraction.
            const string extractedPath = "extracted.bmp";
            using (FileStream fileOut = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
            {
                imageStream.CopyTo(fileOut);
            }

            // 6. Validate that the extracted image file exists and is non‑empty.
            if (!File.Exists(extractedPath) || new FileInfo(extractedPath).Length == 0)
                throw new InvalidOperationException("Extracted image file was not created correctly.");
        }
    }
}
