using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Sample content for PDF/UA compliance.");

        // Set PDF save options to PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        // Save the document as a PDF/UA compliant file.
        string outputPath = "output_pdfua.pdf";
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/UA file was not created.");
    }
}
