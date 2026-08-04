using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the input DOC and output PDF files
        const string inputPath = "input.doc";
        const string outputPath = "output.pdf";

        // Create a sample DOC document
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This is a sample document for PDF/A-2b conversion.");

        // Save the sample document as DOC
        sourceDoc.Save(inputPath, SaveFormat.Doc);

        // Load the DOC document
        Document doc = new Document(inputPath);

        // Set PDF save options to PDF/A-2b compliance.
        // Aspose.Words uses PdfCompliance.PdfA2u for PDF/A-2b level.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfA2u
        };

        // Convert and save the document as PDF with the specified compliance level
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF file was not created.");
    }
}
