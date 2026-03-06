using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceExample
{
    class Program
    {
        static void Main()
        {
            // Input Word document path.
            string inputPath = @"C:\Docs\Sample.docx";

            // Output PDF document path.
            string outputPath = @"C:\Docs\Sample.pdf";

            // Desired PDF compliance level.
            PdfCompliance targetCompliance = PdfCompliance.PdfA1b;

            // Render the document with the specified compliance.
            RenderDocumentWithCompliance(inputPath, outputPath, targetCompliance);
        }

        /// <summary>
        /// Loads a Word document, sets the PDF compliance level, and saves it as PDF.
        /// </summary>
        /// <param name="inputFile">Path to the source .docx/.doc file.</param>
        /// <param name="outputFile">Path where the PDF will be saved.</param>
        /// <param name="compliance">The PDF compliance level to apply.</param>
        static void RenderDocumentWithCompliance(string inputFile, string outputFile, PdfCompliance compliance)
        {
            // Load the document using the Document constructor (lifecycle rule).
            Document doc = new Document(inputFile);

            // Create a PdfSaveOptions object (lifecycle rule).
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // Set the desired compliance level (feature rule).
            saveOptions.Compliance = compliance;

            // Save the document as PDF using the Save method with SaveOptions (lifecycle rule).
            doc.Save(outputFile, saveOptions);
        }
    }
}
