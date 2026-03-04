using System;
using Aspose.Words;

namespace PdfConversionExample
{
    public class PdfConverter
    {
        /// <summary>
        /// Converts a PDF file to DOCX and optionally to other formats.
        /// </summary>
        /// <param name="pdfFilePath">Full path to the source PDF file.</param>
        /// <param name="docxOutputPath">Full path where the DOCX file will be saved.</param>
        /// <param name="htmlOutputPath">Optional path for HTML output (pass null if not needed).</param>
        /// <param name="txtOutputPath">Optional path for plain‑text output (pass null if not needed).</param>
        public void ConvertPdf(string pdfFilePath, string docxOutputPath, string htmlOutputPath = null, string txtOutputPath = null)
        {
            // Load the PDF document. The constructor automatically detects the format.
            Document pdfDocument = new Document(pdfFilePath);

            // Save as DOCX. The file extension determines the SaveFormat.
            pdfDocument.Save(docxOutputPath);

            // If an HTML path is provided, save the document as HTML.
            if (!string.IsNullOrEmpty(htmlOutputPath))
            {
                pdfDocument.Save(htmlOutputPath, SaveFormat.Html);
            }

            // If a plain‑text path is provided, save the document as plain text.
            if (!string.IsNullOrEmpty(txtOutputPath))
            {
                pdfDocument.Save(txtOutputPath, SaveFormat.Text);
            }
        }
    }

    // Example usage
    class Program
    {
        static void Main()
        {
            var converter = new PdfConverter();

            string pdfPath = @"C:\Input\sample.pdf";
            string docxPath = @"C:\Output\sample.docx";
            string htmlPath = @"C:\Output\sample.html";
            string txtPath = @"C:\Output\sample.txt";

            converter.ConvertPdf(pdfPath, docxPath, htmlPath, txtPath);

            Console.WriteLine("Conversion completed.");
        }
    }
}
