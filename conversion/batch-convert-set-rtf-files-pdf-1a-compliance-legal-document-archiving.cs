using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace BatchRtfToPdfA
{
    class Program
    {
        static void Main()
        {
            // Folder containing the source RTF files.
            string sourceFolder = @"C:\LegalDocs\Rtf";
            // Folder where the PDF/A‑1a files will be written.
            string targetFolder = @"C:\LegalDocs\PdfA";

            // Ensure the target folder exists.
            Directory.CreateDirectory(targetFolder);

            // Get all *.rtf files in the source folder (non‑recursive).
            string[] rtfFiles = Directory.GetFiles(sourceFolder, "*.rtf", SearchOption.TopDirectoryOnly);

            // Prepare the PDF save options once – only the compliance level is required.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Set the PDF/A‑1a compliance level for legal archiving.
                Compliance = PdfCompliance.PdfA1a
            };

            foreach (string rtfPath in rtfFiles)
            {
                // Load the RTF document with default load options.
                RtfLoadOptions loadOptions = new RtfLoadOptions();
                Document doc = new Document(rtfPath, loadOptions);

                // Build the output PDF file name (same base name, .pdf extension).
                string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
                string pdfPath = Path.Combine(targetFolder, pdfFileName);

                // Save the document as PDF/A‑1a.
                doc.Save(pdfPath, pdfOptions);
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
