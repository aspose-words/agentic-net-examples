using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsAccessibilityDemo
{
    // Collects warnings that occur during loading or saving.
    public class AccessibilityWarningCollector : IWarningCallback
    {
        // Use a List<T> because WarningInfoCollection does not expose an Add method.
        public List<WarningInfo> Warnings { get; } = new List<WarningInfo>();

        public void Warning(WarningInfo info)
        {
            // Store every warning; you can filter by WarningType if needed.
            Warnings.Add(info);
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source document.
            const string inputPath = @"Input\SampleDocument.docx";

            // Load the document.
            Document doc = new Document(inputPath);

            // Set up the warning collector before any operation that may generate warnings.
            var warningCollector = new AccessibilityWarningCollector();
            doc.WarningCallback = warningCollector;

            // ---------- Save as PDF/A (e.g., PDF/A-4) ----------
            var pdfAOptions = new PdfSaveOptions
            {
                // Choose the desired PDF/A compliance level.
                Compliance = PdfCompliance.PdfA4
            };

            const string pdfAPath = @"Output\SampleDocument_PdfA4.pdf";
            doc.Save(pdfAPath, pdfAOptions);

            // Output warnings that occurred during PDF/A conversion.
            Console.WriteLine("Warnings while saving to PDF/A:");
            foreach (WarningInfo warning in warningCollector.Warnings)
            {
                Console.WriteLine($"- [{warning.Source}] {warning.Description}");
            }

            // Clear previous warnings before the next save.
            warningCollector.Warnings.Clear();

            // ---------- Save as PDF/UA (e.g., PDF/UA-2) ----------
            var pdfUAOptions = new PdfSaveOptions
            {
                // Choose the desired PDF/UA compliance level.
                Compliance = PdfCompliance.PdfUa2
            };

            const string pdfUAPath = @"Output\SampleDocument_PdfUA2.pdf";
            doc.Save(pdfUAPath, pdfUAOptions);

            // Output warnings that occurred during PDF/UA conversion.
            Console.WriteLine("\nWarnings while saving to PDF/UA:");
            foreach (WarningInfo warning in warningCollector.Warnings)
            {
                Console.WriteLine($"- [{warning.Source}] {warning.Description}");
            }
        }
    }
}
