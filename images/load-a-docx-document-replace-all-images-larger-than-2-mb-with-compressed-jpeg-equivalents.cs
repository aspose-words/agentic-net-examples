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
        // File names used in the example.
        const string largeBmpPath = "largeImage.bmp";   // BMP to guarantee large size.
        const string smallJpegPath = "smallImage.jpg";
        const string inputDocPath = "input.docx";
        const string outputDocPath = "output.docx";

        // -----------------------------------------------------------------
        // 1. Create a large BMP image (> 2 MB) using Aspose.Drawing.
        // -----------------------------------------------------------------
        const int largeWidth = 3000;
        const int largeHeight = 3000;
        using (Aspose.Drawing.Bitmap largeBitmap = new Aspose.Drawing.Bitmap(largeWidth, largeHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(largeBitmap))
        {
            // Fill with a solid color and draw a rectangle to make the image deterministic.
            g.Clear(Aspose.Drawing.Color.White);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Blue, 5))
            {
                g.DrawRectangle(pen, 0, 0, largeWidth - 1, largeHeight - 1);
            }

            // Save as BMP (uncompressed) to ensure the file exceeds 2 MB.
            largeBitmap.Save(largeBmpPath, ImageFormat.Bmp);
        }

        // Verify that the generated image is indeed larger than 2 MB.
        if (new FileInfo(largeBmpPath).Length <= 2 * 1024 * 1024)
            throw new Exception("Failed to generate a large image exceeding 2 MB.");

        // -----------------------------------------------------------------
        // 2. Create a small JPEG image that will be used as the compressed replacement.
        // -----------------------------------------------------------------
        const int smallWidth = 200;
        const int smallHeight = 200;
        using (Aspose.Drawing.Bitmap smallBitmap = new Aspose.Drawing.Bitmap(smallWidth, smallHeight))
        using (Aspose.Drawing.Graphics g = Aspose.Drawing.Graphics.FromImage(smallBitmap))
        {
            g.Clear(Aspose.Drawing.Color.LightGray);
            using (Aspose.Drawing.Pen pen = new Aspose.Drawing.Pen(Aspose.Drawing.Color.Red, 3))
            {
                g.DrawEllipse(pen, 10, 10, smallWidth - 20, smallHeight - 20);
            }

            // Save as JPEG with default quality (reasonable compression).
            smallBitmap.Save(smallJpegPath, ImageFormat.Jpeg);
        }

        // -----------------------------------------------------------------
        // 3. Build a sample DOCX containing the large image.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(largeBmpPath);
        doc.Save(inputDocPath);

        // -----------------------------------------------------------------
        // 4. Load the document and replace images larger than 2 MB.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(inputDocPath);
        NodeCollection shapeNodes = loadedDoc.GetChildNodes(NodeType.Shape, true);

        foreach (Shape shape in shapeNodes.OfType<Shape>())
        {
            if (!shape.HasImage)
                continue;

            // Determine the size of the current image.
            using (MemoryStream imgStream = new MemoryStream())
            {
                shape.ImageData.Save(imgStream);
                long imageSize = imgStream.Length;

                // If the image exceeds 2 MB, replace it with the compressed JPEG.
                if (imageSize > 2 * 1024 * 1024)
                {
                    byte[] jpegBytes = File.ReadAllBytes(smallJpegPath);
                    using (MemoryStream ms = new MemoryStream(jpegBytes))
                    {
                        ms.Position = 0; // Ensure the stream is at the beginning.
                        shape.ImageData.SetImage(ms);
                    }
                }
            }
        }

        // -----------------------------------------------------------------
        // 5. Save the modified document.
        // -----------------------------------------------------------------
        loadedDoc.Save(outputDocPath);

        // -----------------------------------------------------------------
        // 6. Validation – ensure the output file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputDocPath) || new FileInfo(outputDocPath).Length == 0)
            throw new Exception("Output document was not created successfully.");

        // Optional cleanup (comment out if you need to inspect the files).
        // File.Delete(largeBmpPath);
        // File.Delete(smallJpegPath);
        // File.Delete(inputDocPath);
    }
}
