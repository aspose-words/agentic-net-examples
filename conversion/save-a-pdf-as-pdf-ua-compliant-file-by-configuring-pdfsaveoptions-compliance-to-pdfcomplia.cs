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
        builder.Writeln("Sample content for PDF/UA compliance.");

        string inputPath = Path.Combine(Directory.GetCurrentDirectory(), "sample.doc");
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            // Use PDF/UA-1 compliance.
            Compliance = PdfCompliance.PdfUa1
        };

        // Save the document as a PDF/UA compliant file.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "output_ua.pdf");
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF/UA output file was not created.");

        // Optional: clean up intermediate files.
        File.Delete(inputPath);
    }
}
