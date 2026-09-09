using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large Word document (500 pages) to simulate a large PDF source.
        Document largeDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(largeDoc);
        for (int i = 1; i <= 500; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 500)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Save the Word document as a regular PDF file.
        const string sourcePdfPath = "large.pdf";
        largeDoc.Save(sourcePdfPath, SaveFormat.Pdf);

        // Load the large PDF file.
        Document pdfDoc = new Document(sourcePdfPath);

        // Prepare PDF/A‑2b (represented by PdfA2u) save options.
        PdfSaveOptions pdfA2bOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u, // PDF/A‑2b compliance
            MemoryOptimization = true
        };

        // Stream the PDF/A‑2b output directly to a MemoryStream.
        using (MemoryStream outputStream = new MemoryStream())
        {
            pdfDoc.Save(outputStream, pdfA2bOptions);

            // Verify that data was written to the stream.
            if (outputStream.Length == 0)
                throw new InvalidOperationException("The output MemoryStream is empty after saving PDF/A‑2b.");

            // Reset the position before any further reading.
            outputStream.Position = 0;

            // For demonstration, write the stream to a file to confirm the result.
            const string resultPdfPath = "result_pdfa2b.pdf";
            using (FileStream fileStream = new FileStream(resultPdfPath, FileMode.Create, FileAccess.Write))
            {
                outputStream.CopyTo(fileStream);
            }

            // Verify that the result file was created.
            if (!File.Exists(resultPdfPath) || new FileInfo(resultPdfPath).Length == 0)
                throw new InvalidOperationException("The PDF/A‑2b file was not created correctly.");
        }

        // Clean up temporary source PDF.
        if (File.Exists(sourcePdfPath))
            File.Delete(sourcePdfPath);
    }
}
