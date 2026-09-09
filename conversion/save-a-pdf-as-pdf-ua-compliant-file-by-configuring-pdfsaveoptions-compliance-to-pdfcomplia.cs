using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a sample Word document.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);
        builder.Writeln("This document will be saved as a PDF/UA compliant file.");

        // Save the sample document locally (bootstrap step).
        const string inputPath = "input.docx";
        source.Save(inputPath, SaveFormat.Docx);

        // Load the saved document.
        Document doc = new Document(inputPath);

        // Configure PDF save options for PDF/UA compliance.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1 // PDF/UA-1 compliance.
        };

        // Save the document as a PDF with the specified compliance.
        const string outputPath = "output.pdf";
        doc.Save(outputPath, pdfOptions);

        // Verify that the PDF file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The PDF/UA compliant file was not created.");

        // Optional: clean up the intermediate DOCX file.
        if (File.Exists(inputPath))
            File.Delete(inputPath);
    }
}
