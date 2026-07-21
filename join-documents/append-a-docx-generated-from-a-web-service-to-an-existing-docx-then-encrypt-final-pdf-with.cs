using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Paths for intermediate and final files.
        string existingDocPath = Path.Combine(outputDir, "Existing.docx");
        string webDocPath = Path.Combine(outputDir, "WebGenerated.docx");
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        string finalPdfPath = Path.Combine(outputDir, "Final.pdf");

        // 1. Create a local DOCX that will serve as the base document.
        Document existingDoc = new Document();
        DocumentBuilder baseBuilder = new DocumentBuilder(existingDoc);
        baseBuilder.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath);

        // 2. Simulate obtaining a DOCX from a web service by creating a local sample document.
        Document webDoc = new Document();
        DocumentBuilder webBuilder = new DocumentBuilder(webDoc);
        webBuilder.Writeln("This content simulates a document retrieved from a web service.");
        webDoc.Save(webDocPath);

        // 3. Append the simulated web‑service document to the existing document.
        Document mergedDoc = new Document(existingDocPath);
        mergedDoc.AppendDocument(webDoc, ImportFormatMode.KeepSourceFormatting);
        mergedDoc.Save(mergedDocPath);

        // 4. Convert the merged document to PDF and encrypt it with a password.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails("UserPassword", "OwnerPassword")
        };
        mergedDoc.Save(finalPdfPath, pdfOptions);

        // 5. Validate that the encrypted PDF was created.
        if (!File.Exists(finalPdfPath))
        {
            throw new InvalidOperationException("The final encrypted PDF was not created.");
        }

        // Program completes without requiring any user interaction.
    }
}
