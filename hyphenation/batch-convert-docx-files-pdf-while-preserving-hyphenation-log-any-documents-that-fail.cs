using System;
using System.IO;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchDocxToPdf
{
    class Program
    {
        static void Main(string[] args)
        {
            // Input folder containing DOCX files
            string inputFolder = @"C:\Docs\Input";
            // Output folder for generated PDFs
            string outputFolder = @"C:\Docs\Output";

            // Ensure the output directory exists
            Directory.CreateDirectory(outputFolder);

            // List to collect information about files that could not be processed
            List<string> failedFiles = new List<string>();

            // Enumerate all .docx files in the input folder (non‑recursive)
            foreach (string docxPath in Directory.GetFiles(inputFolder, "*.docx"))
            {
                try
                {
                    // Load the DOCX document (uses the Document(string) constructor)
                    Document doc = new Document(docxPath);

                    // Enable automatic hyphenation so that the layout (and thus the PDF) preserves hyphenation
                    doc.HyphenationOptions.AutoHyphenation = true;

                    // Build the output PDF file name
                    string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                    string pdfPath = Path.Combine(outputFolder, pdfFileName);

                    // Save the document as PDF (uses Save(string, SaveFormat) overload)
                    doc.Save(pdfPath, SaveFormat.Pdf);
                }
                catch (Exception ex)
                {
                    // Record the failure with the exception message
                    failedFiles.Add($"{Path.GetFileName(docxPath)}: {ex.Message}");
                }
            }

            // Report any failures
            if (failedFiles.Count > 0)
            {
                Console.WriteLine("The following files could not be converted:");
                foreach (string entry in failedFiles)
                {
                    Console.WriteLine(entry);
                }
            }
            else
            {
                Console.WriteLine("All files were successfully converted.");
            }
        }
    }
}
