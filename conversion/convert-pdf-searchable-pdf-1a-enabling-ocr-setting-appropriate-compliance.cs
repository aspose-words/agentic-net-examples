using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

namespace PdfAConversionExample
{
    class Program
    {
        static void Main()
        {
            // Use the current directory for input and output files.
            string baseDir = Directory.GetCurrentDirectory();
            string inputPdfPath = Path.Combine(baseDir, "source.pdf");
            string outputPdfPath = Path.Combine(baseDir, "searchable_a1a.pdf");

            // Ensure a source PDF exists. If not, create a simple one.
            if (!File.Exists(inputPdfPath))
            {
                // Create a simple Word document with some text.
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                builder.Writeln("This is a sample document used for PDF/A‑1a conversion.");

                // Save it as a regular PDF.
                doc.Save(inputPdfPath, SaveFormat.Pdf);
            }

            // Load the PDF document. PdfLoadOptions can be used to customize loading,
            // but for this scenario the default options are sufficient.
            Document pdfDocument = new Document(inputPdfPath, new PdfLoadOptions());

            // Configure PDF save options to produce a PDF/A‑1a compliant file.
            PdfSaveOptions saveOptions = new PdfSaveOptions
            {
                // PDF/A‑1a requires the document structure (tags) to be present,
                // which makes the content searchable.
                Compliance = PdfCompliance.PdfA1a
            };

            // Save the document as PDF/A‑1a. Aspose.Words will perform OCR on any
            // raster images in the PDF during the save operation, embedding the
            // recognized text so that the resulting file is searchable.
            pdfDocument.Save(outputPdfPath, saveOptions);

            Console.WriteLine($"PDF/A‑1a file saved to: {outputPdfPath}");
        }
    }
}
