using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToImages
{
    static void Main()
    {
        // Create a simple PDF document if it does not exist.
        string pdfPath = Path.Combine(Directory.GetCurrentDirectory(), "source.pdf");
        if (!File.Exists(pdfPath))
        {
            var doc = new Document();
            var builder = new DocumentBuilder(doc);
            builder.Writeln("Sample PDF generated for image conversion.");
            doc.Save(pdfPath, SaveFormat.Pdf);
        }

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Convert all pages to JPEG images.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg)
        {
            PageSet = new PageSet(0, pdfDoc.PageCount - 1) // all pages
        };
        string jpegOutput = Path.Combine(Directory.GetCurrentDirectory(), "output.jpg");
        pdfDoc.Save(jpegOutput, jpegOptions);

        // Convert all pages to PNG images.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0, pdfDoc.PageCount - 1) // all pages
        };
        string pngOutput = Path.Combine(Directory.GetCurrentDirectory(), "output.png");
        pdfDoc.Save(pngOutput, pngOptions);

        Console.WriteLine($"PDF converted to images:\n{jpegOutput}\n{pngOutput}");
    }
}
