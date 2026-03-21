using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchDocxToPdf
{
    class Program
    {
        static void Main()
        {
            // Use folders relative to the executable location.
            string baseDir = AppContext.BaseDirectory;
            string sourceFolder = Path.Combine(baseDir, "Input");
            string outputFolder = Path.Combine(baseDir, "Output");

            // Ensure both directories exist.
            Directory.CreateDirectory(sourceFolder);
            Directory.CreateDirectory(outputFolder);

            // Retrieve all DOCX files in the source folder (non‑recursive).
            string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

            if (docxFiles.Length == 0)
            {
                Console.WriteLine($"No DOCX files found in '{sourceFolder}'. Place files there and rerun the program.");
                return;
            }

            foreach (string docxPath in docxFiles)
            {
                // Load the existing DOCX document.
                Document doc = new Document(docxPath);

                // Insert a company‑wide header into each section of the document.
                foreach (Section section in doc.Sections)
                {
                    // Ensure the section has a primary header; create one if missing.
                    HeaderFooter header = section.HeadersFooters[HeaderFooterType.HeaderPrimary];
                    if (header == null)
                    {
                        header = new HeaderFooter(doc, HeaderFooterType.HeaderPrimary);
                        section.HeadersFooters.Add(header);
                    }

                    // Use DocumentBuilder to write the header text.
                    DocumentBuilder builder = new DocumentBuilder(doc);
                    builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
                    builder.Writeln("Company Confidential – Confidential Information");
                }

                // Determine the output PDF file name.
                string pdfFileName = Path.GetFileNameWithoutExtension(docxPath) + ".pdf";
                string pdfPath = Path.Combine(outputFolder, pdfFileName);

                // Save the document as PDF.
                doc.Save(pdfPath, SaveFormat.Pdf);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
