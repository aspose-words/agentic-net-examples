using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfComplianceDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Word document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Load the document from the file system.
            Document doc = new Document(inputPath);

            // ------------------------------------------------------------
            // Example 1: Convert to PDF/A-4 (ISO 19005-4) with PDF/UA-2 compliance.
            // ------------------------------------------------------------
            PdfSaveOptions pdfA4Ua2Options = new PdfSaveOptions
            {
                // PdfA4Ua2 combines PDF/A-4 and PDF/UA-2 standards.
                Compliance = PdfCompliance.PdfA4Ua2
            };

            // Save the document as PDF/A-4 + PDF/UA-2.
            string pdfA4Ua2Path = @"C:\Docs\Output_PdfA4_Ua2.pdf";
            doc.Save(pdfA4Ua2Path, pdfA4Ua2Options);

            // ------------------------------------------------------------
            // Example 2: Convert to PDF/UA-1 (ISO 14289-1) compliance.
            // ------------------------------------------------------------
            PdfSaveOptions pdfUa1Options = new PdfSaveOptions
            {
                // Set compliance to PDF/UA-1.
                Compliance = PdfCompliance.PdfUa1
            };

            // Save the document as PDF/UA-1.
            string pdfUa1Path = @"C:\Docs\Output_PdfUa1.pdf";
            doc.Save(pdfUa1Path, pdfUa1Options);

            // ------------------------------------------------------------
            // Example 3: Convert to PDF/A-1b (ISO 19005-1) compliance.
            // ------------------------------------------------------------
            PdfSaveOptions pdfA1bOptions = new PdfSaveOptions
            {
                // Set compliance to PDF/A-1b.
                Compliance = PdfCompliance.PdfA1b
            };

            // Save the document as PDF/A-1b.
            string pdfA1bPath = @"C:\Docs\Output_PdfA1b.pdf";
            doc.Save(pdfA1bPath, pdfA1bOptions);

            Console.WriteLine("Documents have been saved with the requested PDF compliance levels.");
        }
    }
}
