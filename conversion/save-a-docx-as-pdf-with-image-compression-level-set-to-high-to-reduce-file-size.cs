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
        // Create a sample DOCX document with an image.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOCX with an image.");

        // Generate a simple bitmap using Aspose.Drawing.
        using (Bitmap bitmap = new Bitmap(100, 100))
        {
            // Fill the bitmap with a solid color.
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.Blue);
            }

            // Save the bitmap to a memory stream in PNG format.
            using (MemoryStream imageStream = new MemoryStream())
            {
                bitmap.Save(imageStream, ImageFormat.Png);
                imageStream.Position = 0; // Reset position before inserting.

                // Insert the image into the document.
                builder.InsertImage(imageStream);
            }
        }

        // Save the source document as DOCX.
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the DOCX document.
        Document doc = new Document(inputPath);

        // Configure PDF save options with high image compression.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ImageCompression = PdfImageCompression.Jpeg,
            JpegQuality = 50 // Lower quality results in higher compression.
        };

        // Save the document as PDF using the configured options.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Validate that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
