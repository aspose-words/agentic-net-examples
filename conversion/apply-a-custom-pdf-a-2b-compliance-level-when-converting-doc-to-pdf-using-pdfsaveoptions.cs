using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample DOC file.
        const string inputPath = "sample.doc";
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This is a sample document for PDF/A-2b conversion.");
        source.Save(inputPath, SaveFormat.Doc);

        // Load the DOC file.
        Document doc = new Document(inputPath);

        // Set PDF save options to a PDF/A‑2 compliant level.
        // The enum does not contain a PdfA2b value; PdfA2u is the closest PDF/A‑2 level supported.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Save the document as PDF with the specified compliance level.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created and is not empty.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Expected PDF output file was not created.");

        if (new FileInfo(outputPath).Length == 0)
            throw new InvalidOperationException("The generated PDF file is empty.");
    }
}
