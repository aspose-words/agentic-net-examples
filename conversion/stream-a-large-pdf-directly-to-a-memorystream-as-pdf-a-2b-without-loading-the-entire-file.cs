using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large Word document in memory.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Generate many pages to simulate a large document.
        const int pageCount = 1000;
        for (int i = 0; i < pageCount; i++)
        {
            builder.Writeln($"This is page {i + 1} of a large document.");
            if (i < pageCount - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF/A‑2u save options (PDF/A‑2b is represented by PdfA2u in Aspose.Words).
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u, // PDF/A‑2b compliance
            MemoryOptimization = true
        };

        // Save the document directly to a MemoryStream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, saveOptions);

            // Verify that data was written.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("The PDF/A‑2b stream is empty.");

            // Reset position for any further reading.
            pdfStream.Position = 0;

            // Optional: write the stream to a file to inspect the result.
            const string outputPath = "LargeDocument_PdfA2b.pdf";
            using (FileStream file = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                pdfStream.CopyTo(file);
            }

            // Verify that the file was created.
            if (!File.Exists(outputPath))
                throw new InvalidOperationException("The output PDF/A‑2b file was not created.");
        }
    }
}
