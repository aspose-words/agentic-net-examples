using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large Word document with many pages.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        const int pageCount = 5000; // Adjust to simulate a large document.
        for (int i = 0; i < pageCount; i++)
        {
            builder.Writeln($"This is page {i + 1}");
            if (i < pageCount - 1)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF/A‑2b compliance.
        // Aspose.Words uses PdfCompliance.PdfA2u for PDF/A‑2b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            MemoryOptimization = true
        };

        // Save the document directly to a MemoryStream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, saveOptions);

            // Verify that the stream contains data.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("The resulting PDF stream is empty.");

            // Reset the position for any subsequent read operations.
            pdfStream.Position = 0;

            // Optional: write the stream to a file to demonstrate the output.
            File.WriteAllBytes("OutputPdfA2b.pdf", pdfStream.ToArray());
        }
    }
}
