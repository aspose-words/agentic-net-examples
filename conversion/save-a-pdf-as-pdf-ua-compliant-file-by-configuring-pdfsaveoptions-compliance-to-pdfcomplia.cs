using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "PdfUaCompliant.pdf");

        // Create a simple Word document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("Hello PDF/UA compliant document.");
        builder.Writeln("This document is saved with PDF/UA compliance.");

        // Configure PDF save options for PDF/UA-1 compliance.
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            Compliance = PdfCompliance.PdfUa1
        };

        // Save the document as PDF/UA.
        doc.Save(outputPath, saveOptions);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The PDF/UA file was not created.", outputPath);

        Console.WriteLine($"PDF/UA compliant file saved to: {outputPath}");
    }
}
