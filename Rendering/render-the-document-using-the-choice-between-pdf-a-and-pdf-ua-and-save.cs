using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceDemo
{
    public class PdfRenderer
    {
        /// <summary>
        /// Loads a Word document, applies the specified PDF compliance level (PDF/A or PDF/UA),
        /// and saves the result as a PDF file.
        /// </summary>
        /// <param name="inputPath">Full path to the source .docx/.doc file.</param>
        /// <param name="outputPath">Full path where the resulting PDF will be saved.</param>
        /// <param name="compliance">Desired PDF compliance level (e.g., PdfCompliance.PdfA1b or PdfCompliance.PdfUa1).</param>
        public static void RenderToPdf(string inputPath, string outputPath, PdfCompliance compliance)
        {
            // Load the source document.
            Document doc = new Document(inputPath);

            // Create PDF save options and set the compliance level.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                Compliance = compliance
            };

            // Save the document as PDF using the configured options.
            doc.Save(outputPath, saveOptions);
        }

        // Example usage.
        public static void Main()
        {
            string sourceDoc = @"C:\Docs\SampleDocument.docx";
            string pdfAOutput = @"C:\Docs\SampleDocument_PdfA1b.pdf";
            string pdfUaOutput = @"C:\Docs\SampleDocument_PdfUa1.pdf";

            // Render as PDF/A-1b.
            RenderToPdf(sourceDoc, pdfAOutput, PdfCompliance.PdfA1b);

            // Render as PDF/UA-1.
            RenderToPdf(sourceDoc, pdfUaOutput, PdfCompliance.PdfUa1);
        }
    }
}
