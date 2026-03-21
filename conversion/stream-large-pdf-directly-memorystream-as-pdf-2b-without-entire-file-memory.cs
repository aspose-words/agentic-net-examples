using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfA2bStreamer
{
    static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("Hello, PDF/A‑2b world!");

        // Configure PDF/A‑2b save options.
        var pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u,
            MemoryOptimization = true
        };

        // Save to a MemoryStream.
        using (var outputStream = new MemoryStream())
        {
            doc.Save(outputStream, pdfOptions);

            // Write the result to a file in the current directory.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "LargeDocument_PdfA2b.pdf");
            File.WriteAllBytes(outputPath, outputStream.ToArray());

            Console.WriteLine($"PDF/A‑2b file saved to: {outputPath}");
        }
    }
}
