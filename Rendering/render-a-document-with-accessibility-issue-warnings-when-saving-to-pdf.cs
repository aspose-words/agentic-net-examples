using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace RenderDocumentWithAccessibilityWarnings
{
    class Program
    {
        static void Main()
        {
            // Load the source document.
            Document doc = new Document("InputDocument.docx");

            // Create a warning collector that will receive all warnings generated during saving.
            var warningCallback = new HandleDocumentWarnings();
            doc.WarningCallback = warningCallback;

            // ---------- Save as PDF/A (PDF/A-1b) ----------
            var pdfAOptions = new PdfSaveOptions
            {
                // Set the PDF compliance level to PDF/A-1b.
                Compliance = PdfCompliance.PdfA1b,

                // Export the document structure so that accessibility information is retained.
                ExportDocumentStructure = true
            };

            doc.Save("OutputDocument_PdfA.pdf", pdfAOptions);

            // ---------- Save as PDF/UA ----------
            var pdfUaOptions = new PdfSaveOptions
            {
                // Set the PDF compliance level to PDF/UA-1.
                Compliance = PdfCompliance.PdfUa1,

                // Export the document structure for better accessibility.
                ExportDocumentStructure = true
            };

            doc.Save("OutputDocument_PdfUA.pdf", pdfUaOptions);

            // After saving, you can inspect the collected warnings.
            Console.WriteLine("Warnings collected during PDF/A and PDF/UA conversion:");
            foreach (WarningInfo warning in warningCallback.Warnings)
            {
                Console.WriteLine($"- {warning.WarningType}: {warning.Description}");
            }
        }
    }

    // -----------------------------------------------------------------------------
    // Helper class that implements IWarningCallback to collect warnings.
    // -----------------------------------------------------------------------------
    public class HandleDocumentWarnings : IWarningCallback
    {
        // Collection of all warnings received.
        public WarningInfoCollection Warnings { get; } = new WarningInfoCollection();

        // This method is called by Aspose.Words whenever a warning occurs.
        public void Warning(WarningInfo info)
        {
            // Store the warning for later inspection.
            // WarningInfoCollection does not expose an Add method; use the Warning method instead.
            Warnings.Warning(info);
        }
    }
}
