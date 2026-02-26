using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitter
{
    /// <summary>
    /// Demonstrates how to load a DOCX file, split it into separate pages,
    /// and save each page as an individual PDF document.
    /// </summary>
    public static class SplitDocumentExample
    {
        /// <summary>
        /// Splits the source document into one‑page PDFs.
        /// </summary>
        /// <param name="sourceDocxPath">Full path to the input DOCX file.</param>
        /// <param name="outputFolder">Folder where the PDF parts will be written.</param>
        public static void SplitDocument(string sourceDocxPath, string outputFolder)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the DOCX document from disk.
            Document sourceDoc = new Document(sourceDocxPath);

            // Iterate over each page in the source document.
            for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
            {
                // Extract a single‑page document (pages are zero‑based).
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

                // Build the output PDF file name.
                string pdfPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.pdf");

                // Save the extracted page as PDF. The Save method infers the format from the ".pdf" extension.
                pageDoc.Save(pdfPath);
            }
        }

        // Example usage.
        public static void Main()
        {
            string inputPath = @"C:\Docs\SampleDocument.docx";
            string outputPath = @"C:\Docs\SplitPdfPages";

            SplitDocument(inputPath, outputPath);

            Console.WriteLine("Document split into individual PDF pages successfully.");
        }
    }
}
