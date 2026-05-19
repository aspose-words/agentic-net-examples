using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a large Word document (e.g., 500 pages) to simulate a big source file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        for (int i = 1; i <= 500; i++)
        {
            builder.Writeln($"This is page {i}");
            if (i < 500)
                builder.InsertBreak(BreakType.PageBreak);
        }

        // Configure PDF/A‑2u (closest to PDF/A‑2b) conversion options.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u, // PDF/A‑2b is not a separate enum value; PdfA2u provides PDF/A‑2 compliance.
            MemoryOptimization = true          // Reduces memory usage during saving.
        };

        // Save the document directly into a MemoryStream.
        using (MemoryStream outputStream = new MemoryStream())
        {
            doc.Save(outputStream, pdfOptions);

            // Verify that data was written.
            if (outputStream.Length == 0)
                throw new InvalidOperationException("The PDF/A‑2u stream is empty.");

            // Reset the stream position for any subsequent reads.
            outputStream.Position = 0;

            // Example output: display the size of the generated PDF/A‑2u stream.
            Console.WriteLine($"Generated PDF/A‑2u stream length: {outputStream.Length} bytes");
        }
    }
}
