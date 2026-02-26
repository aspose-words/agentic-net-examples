using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPdfConversion
{
    class Program
    {
        static void Main()
        {
            // Path to the source document (any format supported by Aspose.Words)
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Path where the PDF will be saved
            string outputPath = @"C:\Docs\ConvertedDocument.pdf";

            // Load the document using the provided load rule
            Document doc = new Document(inputPath);

            // Create a SaveOptions object suitable for PDF using the provided factory method
            SaveOptions pdfSaveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);

            // Example configuration: enable memory optimization for large documents
            pdfSaveOptions.MemoryOptimization = true;

            // Save the document as PDF with the configured options using the provided save rule
            doc.Save(outputPath, pdfSaveOptions);
        }
    }
}
