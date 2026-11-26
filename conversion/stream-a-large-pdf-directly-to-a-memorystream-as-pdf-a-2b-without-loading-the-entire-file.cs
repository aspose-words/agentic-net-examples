using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample document with many pages to simulate a large PDF.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        const int pageCount = 1000;
        for (int i = 1; i <= pageCount; i++)
        {
            builder.Writeln($"This is page {i} of a large document.");
            if (i < pageCount)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF/A‑2b compliance options.
        // If the current Aspose.Words version does not contain PdfA2b, fall back to PdfA2a.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = Enum.IsDefined(typeof(PdfCompliance), "PdfA2b")
                ? (PdfCompliance)Enum.Parse(typeof(PdfCompliance), "PdfA2b")
                : PdfCompliance.PdfA2a
        };

        // Stream the PDF/A‑2b (or fallback) output directly to a MemoryStream.
        using MemoryStream outputStream = new MemoryStream();
        source.Save(outputStream, saveOptions);

        // Verify that data was written to the stream.
        if (outputStream.Length == 0)
            throw new InvalidOperationException("The PDF/A output stream is empty.");

        // Reset the position before any further reading.
        outputStream.Position = 0;

        // Write the stream to a file to demonstrate successful creation (optional).
        const string outputPath = "output_pdfa2b.pdf";
        File.WriteAllBytes(outputPath, outputStream.ToArray());

        // Validate that the file was created and is not empty.
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The PDF/A file was not created correctly.");
    }
}
