using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Paths to the source and destination documents.
        const string destinationPath = @"Destination.docx";
        const string sourcePath = @"Source.docx";
        const string outputPdfPath = @"Result.pdf";

        // Load the destination document.
        Document destinationDoc = new Document(destinationPath);

        // Apply read‑only protection (no password required).
        destinationDoc.Protect(ProtectionType.ReadOnly);

        // Load the source document that will be appended.
        Document sourceDoc = new Document(sourcePath);

        // Append the source document to the protected destination document.
        // Keep the source formatting while appending.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Remove the protection so the final document can be saved normally.
        destinationDoc.Unprotect();

        // Save the combined document as PDF.
        destinationDoc.Save(outputPdfPath, SaveFormat.Pdf);
    }
}
