using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many pages to simulate a large file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add 500 pages of sample text.
        for (int i = 1; i <= 500; i++)
        {
            builder.Writeln($"Page {i}");
            if (i < 500)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF save options for PDF/A‑2u compliance (PDF/A‑2b is represented by PdfA2u).
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u, // PDF/A‑2b equivalent
            MemoryOptimization = true
        };

        // Save the document directly to a memory stream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, saveOptions);

            // Verify that data was written.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("The PDF/A‑2u stream is empty.");

            // Reset position before reading.
            pdfStream.Position = 0;

            // Example: write the stream to a file to demonstrate the result.
            const string outputPath = "output_pdfa2u.pdf";
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(file);
            }

            // Validate that the output file was created.
            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
                throw new InvalidOperationException("Failed to create the output PDF/A‑2u file.");
        }
    }
}
