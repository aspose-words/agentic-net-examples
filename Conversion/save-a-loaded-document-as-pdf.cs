using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source document (any format supported by Aspose.Words).
            string inputPath = @"C:\Docs\input.docx";

            // Path where the PDF will be saved.
            string outputPath = @"C:\Docs\output.pdf";

            // Load the document from the file system.
            Document doc = new Document(inputPath);

            // Create PDF save options (optional – can be omitted if default options are sufficient).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set the PDF to open in full‑screen mode.
                // PageMode = PdfPageMode.FullScreen
            };

            // Save the loaded document as PDF using the specified options.
            doc.Save(outputPath, pdfOptions);

            Console.WriteLine("Document saved as PDF successfully.");
        }
    }
}
