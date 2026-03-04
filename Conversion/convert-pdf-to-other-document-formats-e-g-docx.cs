using System;
using Aspose.Words;

namespace AsposeWordsExamples
{
    public static class PdfConverter
    {
        /// <summary>
        /// Converts a PDF file to the specified Word format (DOCX by default).
        /// </summary>
        /// <param name="pdfPath">Full path to the source PDF file.</param>
        /// <param name="outputPath">Full path where the converted document will be saved.</param>
        /// <param name="format">Desired output format (e.g., SaveFormat.Docx, SaveFormat.Doc, SaveFormat.Rtf, etc.).</param>
        public static void ConvertPdfToWord(string pdfPath, string outputPath, SaveFormat format = SaveFormat.Docx)
        {
            // Load the PDF document using the Document constructor that accepts a file name.
            Document pdfDocument = new Document(pdfPath);

            // Save the loaded document in the requested format.
            pdfDocument.Save(outputPath, format);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Example paths – adjust them to your environment.
            string inputPdf = "C:\\Input\\sample.pdf";
            string outputDocx = "C:\\Output\\sample.docx";

            // Perform conversion.
            PdfConverter.ConvertPdfToWord(inputPdf, outputDocx, SaveFormat.Docx);

            Console.WriteLine($"PDF converted to DOCX successfully: {outputDocx}");
        }
    }
}
