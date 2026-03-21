using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a simple document in memory.
        Document doc = new Document();
        var builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample document generated at runtime.");
        builder.Writeln("It demonstrates streaming conversion to PDF with minimal memory usage.");

        // Define the output PDF path (relative to the executable's folder).
        string outputFile = Path.Combine(AppContext.BaseDirectory, "LargeDocument.pdf");

        // Ensure the output directory exists.
        Directory.CreateDirectory(Path.GetDirectoryName(outputFile)!);

        // Configure PDF save options with memory optimization.
        SaveOptions pdfOptions = SaveOptions.CreateSaveOptions(SaveFormat.Pdf);
        pdfOptions.MemoryOptimization = true;

        // Save the document to a write‑only stream, which writes the PDF incrementally.
        using (FileStream outputStream = File.Create(outputFile))
        {
            doc.Save(outputStream, pdfOptions);
        }

        Console.WriteLine($"PDF saved to: {outputFile}");
    }
}
