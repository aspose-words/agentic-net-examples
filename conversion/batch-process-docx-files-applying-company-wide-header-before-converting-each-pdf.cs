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
            // Folder containing the source DOCX files.
            string sourceFolder = @"C:\Docs\Input";

            // Folder where the resulting PDF files will be written.
            string outputFolder = @"C:\Docs\Output";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Retrieve all DOCX files in the source folder (non‑recursive).
            string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx", SearchOption.TopDirectoryOnly);

            foreach (string docxPath in docxFiles)
            {
                // Load the existing DOCX document.
                Document doc = new Document(docxPath);

                // Insert a company‑wide header into each section of the document.
                // The header will contain the same text for all sections.
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

                // Save the document as PDF using the Save(string, SaveFormat) overload.
                doc.Save(pdfPath, SaveFormat.Pdf);
            }

            Console.WriteLine("Batch processing completed.");
        }
    }
}
