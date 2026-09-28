using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large sample document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        const int pageCount = 500; // Simulate a large document.
        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i} of a large document.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF/A‑2b save options.
        // If the used Aspose.Words version does not contain PdfA2b, fall back to PdfA1b.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Uncomment the line below when PdfA2b is available in the referenced Aspose.Words version.
            // Compliance = PdfCompliance.PdfA2b,
            Compliance = PdfCompliance.PdfA1b // Fallback for older versions.
        };

        // Save directly to a MemoryStream without intermediate files.
        using MemoryStream pdfAStream = new MemoryStream();
        source.Save(pdfAStream, saveOptions);

        // Reset position for any subsequent reading.
        pdfAStream.Position = 0;

        // Validate that the stream contains data.
        if (pdfAStream.Length == 0)
            throw new InvalidOperationException("No PDF/A data was written to the MemoryStream.");

        // The MemoryStream now holds the PDF/A document and can be used further.
        Console.WriteLine($"PDF/A stream length: {pdfAStream.Length} bytes");
    }
}
