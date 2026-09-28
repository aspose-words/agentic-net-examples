using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("Sample DOC content for PDF/A‑1b conversion.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Configure PDF/A‑1b compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA1b
        };

        // Save as PDF/A‑1b.
        doc.Save("output.pdf", pdfOptions);

        // Verify that the output file was created.
        if (!File.Exists("output.pdf"))
            throw new InvalidOperationException("Expected output PDF/A‑1b file was not created.");
    }
}
