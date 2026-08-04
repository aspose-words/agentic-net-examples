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

        // Set up PDF save options to enforce PDF/UA-1 compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        string outputPath = "output_pdfua.pdf";

        // Save the document as a PDF using the configured options.
        doc.Save(outputPath, saveOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException("The PDF/UA file was not created.");
        }
    }
}
