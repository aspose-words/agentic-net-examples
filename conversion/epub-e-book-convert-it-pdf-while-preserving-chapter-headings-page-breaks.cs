using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace EpubToPdfConverter
{
    public static class Converter
    {
        /// <summary>
        /// Loads an EPUB file and saves it as a PDF while keeping chapter headings (as outline entries)
        /// and page breaks intact.
        /// </summary>
        /// <param name="epubFilePath">Full path to the source EPUB file.</param>
        /// <param name="pdfFilePath">Full path where the resulting PDF will be written.</param>
        public static void ConvertEpubToPdf(string epubFilePath, string pdfFilePath)
        {
            // Load the EPUB document. The Document constructor handles format detection.
            Document doc = new Document(epubFilePath);

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Include all heading levels (Heading 1‑9) in the PDF outline (bookmarks).
                OutlineOptions = { HeadingsOutlineLevels = 9 },
                // Export the document structure for better accessibility.
                ExportDocumentStructure = true
            };

            // Save the document as PDF using the configured options.
            doc.Save(pdfFilePath, pdfOptions);
        }
    }

    public class Program
    {
        /// <summary>
        /// Entry point required by the .NET runtime. Calls the converter with command‑line arguments.
        /// </summary>
        /// <param name="args">[0] – path to the source EPUB file, [1] – path to the output PDF file.</param>
        public static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: EpubToPdfConverter <input.epub> <output.pdf>");
                return;
            }

            try
            {
                Converter.ConvertEpubToPdf(args[0], args[1]);
                Console.WriteLine($"Successfully converted '{args[0]}' to PDF at '{args[1]}'.");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Conversion failed: {ex.Message}");
            }
        }
    }
}
