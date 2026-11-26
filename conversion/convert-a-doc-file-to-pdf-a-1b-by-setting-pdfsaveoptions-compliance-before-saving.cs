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
        builder.Writeln("Sample DOC content.");
        source.Save("input.doc", SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document("input.doc");

        // Configure PDF/A-1b compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions();
        saveOptions.Compliance = PdfCompliance.PdfA1b;

        // Save the document as PDF/A-1b.
        doc.Save("output.pdf", saveOptions);

        // Verify that the output file was created.
        if (!File.Exists("output.pdf"))
        {
            throw new InvalidOperationException("Expected output PDF/A-1b was not created.");
        }
    }
}
