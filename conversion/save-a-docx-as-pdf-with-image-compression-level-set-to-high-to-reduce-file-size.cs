using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Drawing;
using Aspose.Drawing.Imaging;

public class Program
{
    public static void Main()
    {
        // Create a sample DOCX document with some text and an embedded image.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Sample DOCX content with an embedded image.");

        // Create a simple bitmap using Aspose.Drawing.
        using (MemoryStream imageStream = new MemoryStream())
        {
            using (Bitmap bitmap = new Bitmap(100, 100))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.Clear(Color.Blue);
                }

                // Save the bitmap to the memory stream as PNG.
                bitmap.Save(imageStream, ImageFormat.Png);
            }

            // Reset the stream position before inserting the image.
            imageStream.Position = 0;
            builder.InsertImage(imageStream);
        }

        // Save the document as DOCX (input file).
        const string inputPath = "input.docx";
        sourceDoc.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document for conversion.
        Document doc = new Document(inputPath);

        // Configure PDF save options with high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Use JPEG compression for images and set a lower quality to increase compression.
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 50
        };

        // Save the document as PDF with the specified options.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
