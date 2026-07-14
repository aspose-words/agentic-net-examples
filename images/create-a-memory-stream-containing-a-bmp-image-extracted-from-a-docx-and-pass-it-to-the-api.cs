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
        // Deterministic file names.
        const string bmpFileName = "sample.bmp";
        const string docxFileName = "sample.docx";

        // -------------------------------------------------
        // 1. Create a sample BMP image using Aspose.Drawing.
        // -------------------------------------------------
        const int imgWidth = 200;
        const int imgHeight = 100;
        using (Aspose.Drawing.Bitmap bitmap = new Aspose.Drawing.Bitmap(imgWidth, imgHeight))
        {
            using (Aspose.Drawing.Graphics graphics = Aspose.Drawing.Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with white colour.
                graphics.Clear(Aspose.Drawing.Color.White);
            }

            // Save the bitmap as BMP to the file system.
            bitmap.Save(bmpFileName, Aspose.Drawing.Imaging.ImageFormat.Bmp);
        }

        // -------------------------------------------------
        // 2. Create a DOCX and insert the BMP image.
        // -------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the BMP image as a Shape to preserve the original format.
        Shape bmpShape = new Shape(doc, ShapeType.Image);
        bmpShape.ImageData.SetImage(bmpFileName);
        // Append the shape to the current paragraph.
        builder.CurrentParagraph.AppendChild(bmpShape);

        // Save the document.
        doc.Save(docxFileName);

        // -------------------------------------------------
        // 3. Load the document and extract the image.
        // -------------------------------------------------
        Document loadedDoc = new Document(docxFileName);
        Shape imageShape = null;

        // Find the first shape that contains an image.
        foreach (Shape shape in loadedDoc.GetChildNodes(NodeType.Shape, true))
        {
            if (shape.HasImage)
            {
                imageShape = shape;
                break;
            }
        }

        if (imageShape == null)
            throw new InvalidOperationException("No image found in the document.");

        // -------------------------------------------------
        // 4. Save the image to a memory stream.
        // -------------------------------------------------
        using (MemoryStream imageStream = new MemoryStream())
        {
            // Save the image data into the stream.
            imageShape.ImageData.Save(imageStream);

            // Reset the stream position before using it.
            imageStream.Position = 0;

            // -------------------------------------------------
            // 5. Pass the stream to the target API.
            //    (For demonstration, we simply read its length.)
            // -------------------------------------------------
            long length = imageStream.Length;
            Console.WriteLine($"Extracted image stream length: {length} bytes.");

            // Optional: write the stream to a file to verify the output.
            const string extractedBmpPath = "extracted.bmp";
            using (FileStream fileOut = new FileStream(extractedBmpPath, FileMode.Create, FileAccess.Write))
            {
                imageStream.CopyTo(fileOut);
            }

            // Validate that the file was created and is not empty.
            if (!File.Exists(extractedBmpPath) || new FileInfo(extractedBmpPath).Length == 0)
                throw new InvalidOperationException("Failed to write the extracted image to disk.");
        }

        // Clean up temporary files (optional).
        // File.Delete(bmpFileName);
        // File.Delete(docxFileName);
        // File.Delete("extracted.bmp");
    }
}
