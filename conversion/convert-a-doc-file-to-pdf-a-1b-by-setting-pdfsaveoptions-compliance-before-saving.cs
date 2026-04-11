using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample.doc");
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Sample_PdfA1b.pdf");

        // Create a simple DOC file.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PDF/A‑1b conversion!");
        doc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document loadedDoc = new Document(inputPath);

        // Configure PDF/A‑1b compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save as PDF/A‑1b.
        loadedDoc.Save(outputPath, pdfOptions);

        // Verify that the output file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/A‑1b file was not created.");

        Console.WriteLine("Conversion completed successfully.");
    }
}
