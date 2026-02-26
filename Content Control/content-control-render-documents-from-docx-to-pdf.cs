using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    public class DocxToPdfRenderer
    {
        /// <summary>
        /// Loads a DOCX file from the specified path and renders it to a PDF file.
        /// </summary>
        /// <param name="inputDocxPath">Full path to the source DOCX document.</param>
        /// <param name="outputPdfPath">Full path where the resulting PDF will be saved.</param>
        public static void RenderDocxToPdf(string inputDocxPath, string outputPdfPath)
        {
            // Load the existing DOCX document using the Document(string) constructor.
            Document doc = new Document(inputDocxPath);

            // Optionally configure PDF save options (e.g., high‑quality rendering).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                UseHighQualityRendering = true,
                UseAntiAliasing = true
            };

            // Save the document as PDF. The overload Save(string, SaveOptions) follows the
            // provided lifecycle rule (load → save) without inventing custom APIs.
            doc.Save(outputPdfPath, pdfOptions);
        }

        // Example usage.
        public static void Main()
        {
            string sourceDocx = @"C:\Docs\SampleDocument.docx";
            string destinationPdf = @"C:\Docs\SampleDocument.pdf";

            RenderDocxToPdf(sourceDocx, destinationPdf);

            Console.WriteLine("DOCX has been successfully rendered to PDF.");
        }
    }
}
