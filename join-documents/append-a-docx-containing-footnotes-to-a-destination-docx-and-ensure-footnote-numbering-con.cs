using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create the destination DOCX.
        var destDoc = new Document();
        var destBuilder = new DocumentBuilder(destDoc);
        destBuilder.Writeln("This is the first paragraph in the destination document.");
        // Simulated footnote text (actual footnote API may not be available in this environment).
        destBuilder.Writeln("[1] Destination footnote 1.");
        destBuilder.Writeln("Another paragraph in the destination document.");

        const string destPath = "destination.docx";
        destDoc.Save(destPath, SaveFormat.Docx);

        // Create the source DOCX.
        var sourceDoc = new Document();
        var sourceBuilder = new DocumentBuilder(sourceDoc);
        sourceBuilder.Writeln("This is the first paragraph in the source document.");
        // Simulated footnote text.
        sourceBuilder.Writeln("[2] Source footnote A.");
        sourceBuilder.Writeln("Second paragraph in the source document.");
        sourceBuilder.Writeln("[3] Source footnote B.");

        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath, SaveFormat.Docx);

        // Append the source document to the destination document.
        destDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document as PDF.
        const string pdfPath = "merged.pdf";
        destDoc.Save(pdfPath, SaveFormat.Pdf);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException($"Failed to create the merged PDF at '{pdfPath}'.");
        }

        // Optional cleanup (commented out to keep the sample files).
        // File.Delete(destPath);
        // File.Delete(sourcePath);
    }
}
