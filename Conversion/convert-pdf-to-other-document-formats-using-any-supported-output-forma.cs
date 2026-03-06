using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfConversionDemo
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Load the PDF document. The constructor automatically detects the format.
        Document pdfDoc = new Document(pdfPath);

        // Convert PDF to DOCX.
        pdfDoc.Save(@"C:\Output\sample.docx", SaveFormat.Docx);

        // Convert PDF to HTML.
        pdfDoc.Save(@"C:\Output\sample.html", SaveFormat.Html);

        // Convert PDF to RTF.
        pdfDoc.Save(@"C:\Output\sample.rtf", SaveFormat.Rtf);

        // Convert PDF to plain text.
        pdfDoc.Save(@"C:\Output\sample.txt", SaveFormat.Text);

        // Convert PDF to ODT.
        pdfDoc.Save(@"C:\Output\sample.odt", SaveFormat.Odt);

        // Convert PDF to EPUB.
        pdfDoc.Save(@"C:\Output\sample.epub", SaveFormat.Epub);

        // Convert PDF to XPS.
        pdfDoc.Save(@"C:\Output\sample.xps", SaveFormat.Xps);

        // Convert PDF to PNG (first page rendered as an image).
        // Use ImageSaveOptions to specify the image format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        pdfDoc.Save(@"C:\Output\sample.png", pngOptions);

        // Convert PDF to JPEG (first page rendered as an image).
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        pdfDoc.Save(@"C:\Output\sample.jpg", jpegOptions);
    }
}
