using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is saved as PDF/A‑1b.");

        // Configure PDF save options for PDF/A‑1b compliance.
        // Note: Embedding an ICC profile requires the IccProfile property,
        // which may not be available in older Aspose.Words versions.
        // The example demonstrates the standard way to set PDF/A‑1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save the document as PDF/A‑1b.
        const string outputPdf = "output.pdf";
        doc.Save(outputPdf, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("The PDF/A‑1b file was not created.");

        Console.WriteLine($"PDF/A‑1b file '{outputPdf}' has been created successfully.");
    }
}
