using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Drawing;

public class Program
{
    public static void Main()
    {
        // 1. Create a deterministic BMP image using Aspose.Drawing.
        const string bmpPath = "sample.bmp";
        const int width = 100;
        const int height = 100;

        using (Bitmap bitmap = new Bitmap(width, height))
        {
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                // Fill the bitmap with white background.
                graphics.Clear(Color.White);
                // (Optional) Draw a simple rectangle.
                graphics.DrawRectangle(new Pen(Color.Black), 10, 10, width - 20, height - 20);
            }

            // Save the bitmap as BMP – this will be the image we embed in the document.
            bitmap.Save(bmpPath);
        }

        // 2. Create a new Word document and insert the BMP image.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.InsertImage(bmpPath);
        const string docPath = "sample.docx";
        doc.Save(docPath);

        // 3. Load the document (demonstrating load from file) and locate the first shape that contains an image.
        Document loadedDoc = new Document(docPath);
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

        // 4. Save the image data to a memory stream (BMP format is preserved because the source image is BMP).
        using (MemoryStream imageStream = new MemoryStream())
        {
            imageShape.ImageData.Save(imageStream);
            // Reset the stream position before any further read operations.
            imageStream.Position = 0;

            // 5. Pass the memory stream to a sample API method.
            ProcessImageStream(imageStream);
        }

        // Cleanup: ensure the generated files exist.
        if (!File.Exists("extracted.bmp"))
            throw new FileNotFoundException("The extracted image file was not created.");

        Console.WriteLine("Image extraction and processing completed successfully.");
    }

    // Sample API that consumes a stream containing a BMP image.
    private static void ProcessImageStream(Stream imageStream)
    {
        // For demonstration, copy the stream to a file.
        const string outputPath = "extracted.bmp";

        // Ensure the stream is at the beginning.
        if (imageStream.CanSeek)
            imageStream.Position = 0;

        using (FileStream fileStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        {
            imageStream.CopyTo(fileStream);
        }
    }
}
