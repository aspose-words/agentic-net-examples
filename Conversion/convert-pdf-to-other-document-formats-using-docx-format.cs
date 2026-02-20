using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfConversionExample
{
    public class PdfConverter
    {
        /// <summary>
        /// Converts a PDF file to DOCX and then to several other formats.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="outputDirectory">Directory where converted files will be saved.</param>
        public static void ConvertPdfToOtherFormats(string pdfPath, string outputDirectory)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDirectory);

            // ---------- Load PDF ----------
            // Use PdfLoadOptions to control PDF loading behavior if needed.
            var pdfLoadOptions = new PdfLoadOptions();
            Document pdfDocument = new Document(pdfPath, pdfLoadOptions);

            // ---------- Save as DOCX (intermediate) ----------
            string docxPath = Path.Combine(outputDirectory, "Converted.docx");
            var docxSaveOptions = new OoxmlSaveOptions(SaveFormat.Docx);
            pdfDocument.Save(docxPath, docxSaveOptions);

            // ---------- Load the intermediate DOCX ----------
            Document docxDocument = new Document(docxPath);

            // ---------- Save as HTML ----------
            string htmlPath = Path.Combine(outputDirectory, "Converted.html");
            var htmlSaveOptions = new HtmlSaveOptions(SaveFormat.Html);
            docxDocument.Save(htmlPath, htmlSaveOptions);

            // ---------- Save as ODT ----------
            string odtPath = Path.Combine(outputDirectory, "Converted.odt");
            var odtSaveOptions = new OdtSaveOptions(SaveFormat.Odt);
            docxDocument.Save(odtPath, odtSaveOptions);

            // ---------- Save as plain text ----------
            string txtPath = Path.Combine(outputDirectory, "Converted.txt");
            docxDocument.Save(txtPath, SaveFormat.Text);

            // ---------- Save as XPS ----------
            string xpsPath = Path.Combine(outputDirectory, "Converted.xps");
            var xpsSaveOptions = new XpsSaveOptions();
            docxDocument.Save(xpsPath, xpsSaveOptions);

            // Additional formats can be added similarly using the appropriate SaveOptions classes.
        }

        // Example usage
        public static void Main()
        {
            string pdfFile = @"C:\Input\sample.pdf";
            string outputDir = @"C:\Output";

            ConvertPdfToOtherFormats(pdfFile, outputDir);
            Console.WriteLine("Conversion completed.");
        }
    }
}
