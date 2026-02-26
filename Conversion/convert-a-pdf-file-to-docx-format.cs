using System;
using Aspose.Words;

namespace PdfToDocxConverter
{
    public static class Converter
    {
        /// <summary>
        /// Converts a PDF file to DOCX format using Aspose.Words.
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="docxPath">Full path where the resulting DOCX file will be saved.</param>
        public static void ConvertPdfToDocx(string pdfPath, string docxPath)
        {
            // Load the PDF document. The Document constructor automatically detects the format.
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded document as DOCX. The Save method determines the format from the file extension.
            pdfDocument.Save(docxPath);
        }

        // Example usage
        public static void Main()
        {
            // Adjust these paths as needed.
            string sourcePdf = @"C:\Input\sample.pdf";
            string targetDocx = @"C:\Output\sample.docx";

            ConvertPdfToDocx(sourcePdf, targetDocx);

            Console.WriteLine("Conversion completed.");
        }
    }
}
