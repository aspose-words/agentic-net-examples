using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create the destination document in memory.
        var destinationDoc = new Document();
        var destBuilder = new DocumentBuilder(destinationDoc);
        destBuilder.Writeln("This is the destination document.");

        // Apply read‑only protection (no password required).
        destinationDoc.Protect(ProtectionType.ReadOnly);

        // Create the source document in memory.
        var sourceDoc = new Document();
        var srcBuilder = new DocumentBuilder(sourceDoc);
        srcBuilder.Writeln("This is the source document.");

        // Append the source document to the protected destination document,
        // keeping the source formatting.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Remove protection so the final document can be saved normally.
        destinationDoc.Unprotect();

        // Save the combined document as PDF.
        const string outputPdfPath = "Result.pdf";
        destinationDoc.Save(outputPdfPath, SaveFormat.Pdf);

        Console.WriteLine($"PDF saved to {outputPdfPath}");
    }
}
