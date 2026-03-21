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
            // Use folders relative to the executable location.
            string baseDir = AppContext.BaseDirectory;
            string sourceFolder = Path.Combine(baseDir, "Rtf");
            string targetFolder = Path.Combine(baseDir, "PdfA");

            // Ensure both folders exist.
            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(targetFolder);

            // If there are no RTF files, create a simple sample file.
            if (Directory.GetFiles(sourceFolder, "*.rtf").Length == 0)
            {
                string samplePath = Path.Combine(sourceFolder, "Sample.rtf");
                File.WriteAllText(samplePath,
                    @"{\rtf1\ansi\deff0{\fonttbl{\f0\fswiss Helvetica;}}\n" +
                    @"\viewkind4\uc1\pard\fs20 This is a sample RTF document.\par}");
            }

            // Get all *.rtf files in the source folder (non‑recursive).
            string[] rtfFiles = Directory.GetFiles(sourceFolder, "*.rtf", SearchOption.TopDirectoryOnly);

            // Prepare the PDF save options – set PDF/A‑1a compliance.
            var pdfOptions = new PdfSaveOptions
            {
                Compliance = PdfCompliance.PdfA1a
            };

            foreach (string rtfPath in rtfFiles)
            {
                // Load the RTF document.
                var loadOptions = new RtfLoadOptions();
                var doc = new Document(rtfPath, loadOptions);

                // Build the output PDF file name.
                string pdfFileName = Path.GetFileNameWithoutExtension(rtfPath) + ".pdf";
                string pdfPath = Path.Combine(targetFolder, pdfFileName);

                // Save as PDF/A‑1a.
                doc.Save(pdfPath, pdfOptions);
            }

            Console.WriteLine("Batch conversion completed.");
            Console.WriteLine($"Source folder: {sourceFolder}");
            Console.WriteLine($"Target folder: {targetFolder}");
        }
    }
}
