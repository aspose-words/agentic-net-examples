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
        builder.Writeln("This is a sample document for PDF/A-2b conversion.");
        string inputPath = "input.doc";
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Configure PDF save options for PDF/A-2b compliance.
        // The PdfCompliance enum may not expose PdfA2b in older library versions,
        // so we set the underlying enum value directly (value 3 corresponds to PDF/A-2b).
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = (PdfCompliance)3 // PDF/A-2b
        };

        // Save as PDF with the specified compliance.
        string outputPath = "output.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("Expected output PDF was not created.");
        }
    }
}
