using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ContentControlPdfRender
{
    public class Renderer
    {
        /// <summary>
        /// Loads a DOCX file that may contain content controls and renders it to a PDF document.
        /// </summary>
        /// <param name="inputPath">Full path to the source DOCX file.</param>
        /// <param name="outputPath">Full path where the resulting PDF will be saved.</param>
        public static void RenderDocxToPdf(string inputPath, string outputPath)
        {
            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Example: render DrawingML shapes as they appear (instead of fallback shapes).
            pdfOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

            // Example: enable high‑quality rendering for better visual fidelity.
            pdfOptions.UseHighQualityRendering = true;

            // Save the document as PDF using the specified options.
            doc.Save(outputPath, pdfOptions);
        }

        // Example usage.
        public static void Main()
        {
            string sourceDocx = @"C:\Docs\SampleWithContentControls.docx";
            string targetPdf = @"C:\Docs\SampleWithContentControls.pdf";

            RenderDocxToPdf(sourceDocx, targetPdf);

            Console.WriteLine("DOCX has been successfully rendered to PDF.");
        }
    }
}
