using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large document with many paragraphs.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 0; i < 5000; i++)
        {
            builder.Writeln($"Paragraph {i + 1}: The quick brown fox jumps over the lazy dog.");
        }

        // Configure PDF/A‑2b (baseline) compliance.
        // In Aspose.Words, PDF/A‑2b is represented by PdfA2u.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            MemoryOptimization = true
        };

        // Save the document directly to a MemoryStream.
        using (MemoryStream pdfStream = new MemoryStream())
        {
            doc.Save(pdfStream, saveOptions);
            pdfStream.Position = 0; // Reset for any subsequent reading.

            // Verify that the stream contains data.
            if (pdfStream.Length == 0)
                throw new InvalidOperationException("The PDF/A‑2b stream is empty.");
        }
    }
}
