using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfConversionExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source PDF file.
            string pdfPath = @"C:\Docs\sample.pdf";

            // Load the PDF document using PdfLoadOptions.
            // PdfLoadOptions allows you to control how the PDF is interpreted.
            var loadOptions = new PdfLoadOptions
            {
                // Example: skip images while loading (set to false to keep them).
                SkipPdfImages = false,
                // Example: set password if the PDF is encrypted.
                Password = null
            };
            Document pdfDocument = new Document(pdfPath, loadOptions);

            // -----------------------------------------------------------------
            // Convert PDF to DOCX (Microsoft Word Open XML format).
            // -----------------------------------------------------------------
            string docxPath = @"C:\Docs\sample_converted.docx";
            pdfDocument.Save(docxPath, SaveFormat.Docx);

            // -----------------------------------------------------------------
            // Convert PDF to HTML.
            // -----------------------------------------------------------------
            string htmlPath = @"C:\Docs\sample_converted.html";
            var htmlOptions = new HtmlSaveOptions
            {
                // Export images as separate files (default behavior).
                ExportImagesAsBase64 = false,
                // Keep the original document structure.
                ExportGeneratorName = true,
                // Use UTF-8 encoding without BOM.
                Encoding = new System.Text.UTF8Encoding(false)
            };
            pdfDocument.Save(htmlPath, htmlOptions);

            // -----------------------------------------------------------------
            // Convert PDF to XPS (XML Paper Specification).
            // -----------------------------------------------------------------
            string xpsPath = @"C:\Docs\sample_converted.xps";
            pdfDocument.Save(xpsPath, SaveFormat.Xps);

            // -----------------------------------------------------------------
            // Convert PDF to plain text.
            // -----------------------------------------------------------------
            string txtPath = @"C:\Docs\sample_converted.txt";
            pdfDocument.Save(txtPath, SaveFormat.Text);

            // -----------------------------------------------------------------
            // Convert PDF to ODT (OpenDocument Text).
            // -----------------------------------------------------------------
            string odtPath = @"C:\Docs\sample_converted.odt";
            pdfDocument.Save(odtPath, SaveFormat.Odt);

            // -----------------------------------------------------------------
            // Optional: Detect the format of the original file using FileFormatUtil.
            // -----------------------------------------------------------------
            FileFormatInfo formatInfo = FileFormatUtil.DetectFileFormat(pdfPath);
            Console.WriteLine($"Original file format: {formatInfo.LoadFormat}");
            Console.WriteLine($"Is encrypted: {formatInfo.IsEncrypted}");
            Console.WriteLine($"Has macros: {formatInfo.HasMacros}");
        }
    }
}
