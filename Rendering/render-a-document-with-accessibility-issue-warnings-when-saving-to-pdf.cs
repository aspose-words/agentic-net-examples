using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceDemo
{
    // Collects warnings generated during document processing.
    public class WarningCollector : IWarningCallback
    {
        // Use a List to store the warnings because WarningInfoCollection is read‑only.
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        // Called by Aspose.Words when a warning occurs.
        public void Warning(WarningInfo info)
        {
            // Store all warnings; you can filter by WarningType if needed.
            Warnings.Add(info);
        }

        // Clears previously collected warnings.
        public void Clear()
        {
            Warnings.Clear();
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source document.
            const string inputPath = @"ArtifactsDir\Input.docx";

            // Load the document.
            Document doc = new Document(inputPath);

            // Set up the warning collector before any processing.
            WarningCollector collector = new WarningCollector();
            doc.WarningCallback = collector;

            // -----------------------------------------------------------------
            // 1. Save as PDF/A-2u (preserves visual appearance and allows text extraction).
            // -----------------------------------------------------------------
            collector.Clear();

            PdfSaveOptions pdfA2uOptions = new PdfSaveOptions
            {
                // Set the compliance level to PDF/A-2u.
                Compliance = PdfCompliance.PdfA2u,

                // Export document structure to help with accessibility (optional, ignored for PDF/A-2u).
                ExportDocumentStructure = true
            };

            // Save the PDF/A-2u file.
            doc.Save(@"ArtifactsDir\Output_PdfA2u.pdf", pdfA2uOptions);

            // Output any warnings that occurred during the save operation.
            Console.WriteLine("Warnings while saving as PDF/A-2u:");
            foreach (WarningInfo warning in collector.Warnings)
            {
                Console.WriteLine($"- [{warning.WarningType}] {warning.Description}");
            }

            // -----------------------------------------------------------------
            // 2. Save as PDF/UA-1 (PDF/Universal Accessibility).
            // -----------------------------------------------------------------
            collector.Clear();

            PdfSaveOptions pdfUa1Options = new PdfSaveOptions
            {
                // Set the compliance level to PDF/UA-1.
                Compliance = PdfCompliance.PdfUa1,

                // Export document structure is required for PDF/UA compliance.
                ExportDocumentStructure = true
            };

            // Save the PDF/UA-1 file.
            doc.Save(@"ArtifactsDir\Output_PdfUa1.pdf", pdfUa1Options);

            // Output any warnings that occurred during the save operation.
            Console.WriteLine("\nWarnings while saving as PDF/UA-1:");
            foreach (WarningInfo warning in collector.Warnings)
            {
                Console.WriteLine($"- [{warning.WarningType}] {warning.Description}");
            }
        }
    }
}
