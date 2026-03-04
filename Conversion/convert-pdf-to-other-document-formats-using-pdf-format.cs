using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace PdfConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Input\sample.pdf";

            // Load the PDF document using the Document constructor (lifecycle rule: load).
            Document pdfDocument = new Document(pdfPath);

            // Convert PDF to DOCX.
            string docxPath = @"C:\Output\sample.docx";
            pdfDocument.Save(docxPath, SaveFormat.Docx); // Save method (lifecycle rule: save).

            // Convert PDF to HTML.
            string htmlPath = @"C:\Output\sample.html";
            pdfDocument.Save(htmlPath, SaveFormat.Html); // Save method with explicit format.

            // Convert PDF to plain text.
            string txtPath = @"C:\Output\sample.txt";
            pdfDocument.Save(txtPath, SaveFormat.Text); // Save method with explicit format.

            // Convert PDF to RTF.
            string rtfPath = @"C:\Output\sample.rtf";
            pdfDocument.Save(rtfPath, SaveFormat.Rtf); // Save method with explicit format.

            // Convert PDF to EPUB.
            string epubPath = @"C:\Output\sample.epub";
            pdfDocument.Save(epubPath, SaveFormat.Epub); // Save method with explicit format.

            // Convert PDF to ODT.
            string odtPath = @"C:\Output\sample.odt";
            pdfDocument.Save(odtPath, SaveFormat.Odt); // Save method with explicit format.

            // Convert PDF to XPS.
            string xpsPath = @"C:\Output\sample.xps";
            pdfDocument.Save(xpsPath, SaveFormat.Xps); // Save method with explicit format.

            // Convert PDF to PNG (image of the first page).
            // Use ImageSaveOptions to specify image format.
            string pngPath = @"C:\Output\sample.png";
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the first page.
                PageSet = new PageSet(0)
            };
            pdfDocument.Save(pngPath, pngOptions); // Save method with SaveOptions.

            Console.WriteLine("PDF conversion completed successfully.");
        }
    }
}
