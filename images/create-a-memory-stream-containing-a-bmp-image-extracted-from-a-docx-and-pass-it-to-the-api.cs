using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;               // Aspose.Drawing.Common namespace
using Aspose.Drawing.Imaging;       // For ImageFormat

public class Program
{
    public static void Main()
    {
        // -----------------------------------------------------------------
        // 1. Create a deterministic BMP image file to be used as input.
        // -----------------------------------------------------------------
        const string bmpPath = "sample.bmp";
        const int width = 100;
        const int height = 100;

        // Create a white bitmap.
        using (Bitmap bitmap = new Bitmap(width, height))
        using (Graphics graphics = Graphics.FromImage(bitmap))
        {
            graphics.Clear(Color.White);
            // Save as BMP.
            bitmap.Save(bmpPath, ImageFormat.Bmp);
        }

        // Verify that the BMP file was created.
        if (!File.Exists(bmpPath))
            throw new FileNotFoundException("Failed to create the sample BMP image.", bmpPath);

        // -----------------------------------------------------------------
        // 2. Create a DOCX document and insert the BMP image.
        // -----------------------------------------------------------------
        const string docPath = "sample.docx";
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert the BMP image into the document.
        builder.InsertImage(bmpPath);

        // Save the document.
        doc.Save(docPath);

        // Verify that the DOCX file was created.
        if (!File.Exists(docPath))
            throw new FileNotFoundException("Failed to create the sample DOCX document.", docPath);

        // -----------------------------------------------------------------
        // 3. Load the document and extract the first image (BMP) into a MemoryStream.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(docPath);

        // Find the first shape that actually contains an image.
        Shape imageShape = null;
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

        // Save the image data to a memory stream.
        using (MemoryStream imageStream = new MemoryStream())
        {
            // The image was originally a BMP, so this will preserve the BMP format.
            imageShape.ImageData.Save(imageStream);

            // Reset the stream position before any further use.
            imageStream.Position = 0;

            // -----------------------------------------------------------------
            // 4. Example: pass the memory stream to an API (placeholder).
            // -----------------------------------------------------------------
            // For demonstration, we simply output the size of the stream.
            Console.WriteLine($"Extracted image stream length: {imageStream.Length} bytes");

            // If you had an API method like: void UploadImage(Stream stream);
            // you would call: UploadImage(imageStream);
        }

        // Cleanup: optional removal of temporary files.
        // File.Delete(bmpPath);
        // File.Delete(docPath);
    }
}
