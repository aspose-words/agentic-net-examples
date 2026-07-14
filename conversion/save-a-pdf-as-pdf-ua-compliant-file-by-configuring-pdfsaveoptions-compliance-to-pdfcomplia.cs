using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a new blank document and add some text.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is PDF/UA compliant.");

        // Set up PDF save options to enforce PDF/UA compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        string outputPath = "output_pdfua.pdf";

        // Save the document as a PDF/UA file.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created successfully.
        if (!File.Exists(outputPath))
        {
            throw new InvalidOperationException($"The PDF/UA file '{outputPath}' was not created.");
        }
    }
}
