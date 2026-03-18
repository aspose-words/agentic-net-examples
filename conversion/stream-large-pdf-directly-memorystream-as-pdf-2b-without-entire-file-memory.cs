using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA2bStreamer
{
    static void Main()
    {
        // Path to the source PDF file (large document)
        const string sourcePdfPath = @"C:\Temp\LargeDocument.pdf";

        // Load the source PDF into an Aspose.Words Document.
        // The Document constructor reads the file from the stream.
        using (FileStream sourceStream = new FileStream(sourcePdfPath, FileMode.Open, FileAccess.Read))
        {
            Document doc = new Document(sourceStream);

            // Create a SaveOptions object suitable for PDF format.
            // The factory method returns a PdfSaveOptions instance.
            SaveOptions saveOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
            PdfSaveOptions pdfOptions = (PdfSaveOptions)saveOptions;

            // Set PDF/A‑2b compliance. In Aspose.Words PdfA2b is represented by PdfA2u.
            pdfOptions.Compliance = PdfCompliance.PdfA2u;

            // Enable memory‑optimised saving to reduce the peak memory usage.
            pdfOptions.MemoryOptimization = true;

            // Stream the result directly into a MemoryStream.
            using (MemoryStream outputStream = new MemoryStream())
            {
                doc.Save(outputStream, pdfOptions);

                // At this point outputStream contains the PDF/A‑2b document.
                // Example: write the stream to a file (optional).
                File.WriteAllBytes(@"C:\Temp\LargeDocument_PdfA2b.pdf", outputStream.ToArray());
            }
        }
    }
}
