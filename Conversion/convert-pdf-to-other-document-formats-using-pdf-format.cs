using System;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfConversionExample
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Load the PDF document using the Document constructor (load rule).
        Document pdfDoc = new Document(pdfPath);

        // -----------------------------------------------------------------
        // Convert PDF to DOCX
        // -----------------------------------------------------------------
        string docxPath = @"C:\Output\sample_converted.docx";
        pdfDoc.Save(docxPath, SaveFormat.Docx); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to HTML
        // -----------------------------------------------------------------
        string htmlPath = @"C:\Output\sample_converted.html";
        pdfDoc.Save(htmlPath, SaveFormat.Html); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to plain text
        // -----------------------------------------------------------------
        string txtPath = @"C:\Output\sample_converted.txt";
        pdfDoc.Save(txtPath, SaveFormat.Text); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to RTF
        // -----------------------------------------------------------------
        string rtfPath = @"C:\Output\sample_converted.rtf";
        pdfDoc.Save(rtfPath, SaveFormat.Rtf); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to EPUB
        // -----------------------------------------------------------------
        string epubPath = @"C:\Output\sample_converted.epub";
        pdfDoc.Save(epubPath, SaveFormat.Epub); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to ODT (OpenDocument Text)
        // -----------------------------------------------------------------
        string odtPath = @"C:\Output\sample_converted.odt";
        pdfDoc.Save(odtPath, SaveFormat.Odt); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to XPS
        // -----------------------------------------------------------------
        string xpsPath = @"C:\Output\sample_converted.xps";
        pdfDoc.Save(xpsPath, SaveFormat.Xps); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to Markdown
        // -----------------------------------------------------------------
        string mdPath = @"C:\Output\sample_converted.md";
        pdfDoc.Save(mdPath, SaveFormat.Markdown); // Save rule with explicit format.

        // -----------------------------------------------------------------
        // Convert PDF to PNG (first page only) using ImageSaveOptions.
        // -----------------------------------------------------------------
        string pngPath = @"C:\Output\sample_converted.png";
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Render only the first page.
            PageSet = new PageSet(0)
        };
        pdfDoc.Save(pngPath, pngOptions); // Save rule with SaveOptions.

        // -----------------------------------------------------------------
        // Convert PDF to TIFF (all pages) using ImageSaveOptions.
        // -----------------------------------------------------------------
        string tiffPath = @"C:\Output\sample_converted.tiff";
        ImageSaveOptions tiffOptions = new ImageSaveOptions(SaveFormat.Tiff);
        pdfDoc.Save(tiffPath, tiffOptions); // Save rule with SaveOptions.

        // All conversions completed.
        Console.WriteLine("PDF conversion to multiple formats completed successfully.");
    }
}
