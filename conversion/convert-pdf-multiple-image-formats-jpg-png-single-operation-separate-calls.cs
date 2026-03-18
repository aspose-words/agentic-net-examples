using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToImages
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Docs\source.pdf";

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // ---------- Convert to JPEG ----------
        // Create ImageSaveOptions for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        // Optional: set resolution, quality, etc. (defaults are acceptable for basic conversion).

        // Save the first page of the PDF as a JPEG image.
        // If you need all pages, iterate over PageSet as shown in the documentation.
        pdfDoc.Save(@"C:\Docs\output.jpg", jpegOptions);

        // ---------- Convert to PNG ----------
        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

        // Save the first page of the PDF as a PNG image.
        pdfDoc.Save(@"C:\Docs\output.png", pngOptions);
    }
}
