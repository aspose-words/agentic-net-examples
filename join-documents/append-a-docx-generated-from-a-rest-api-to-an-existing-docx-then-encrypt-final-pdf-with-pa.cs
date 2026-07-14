using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare an output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Define file paths.
        string existingDocPath = Path.Combine(outputDir, "Existing.docx");
        string apiDocPath = Path.Combine(outputDir, "ApiGenerated.docx");
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        string finalPdfPath = Path.Combine(outputDir, "FinalEncrypted.pdf");

        // -----------------------------------------------------------------
        // 1. Create a local DOCX that represents an existing document.
        // -----------------------------------------------------------------
        Document existingDoc = new Document();
        DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);
        existingBuilder.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath);

        // -----------------------------------------------------------------
        // 2. Simulate a DOCX generated from a REST API.
        // -----------------------------------------------------------------
        Document apiDoc = new Document();
        DocumentBuilder apiBuilder = new DocumentBuilder(apiDoc);
        apiBuilder.Writeln("Content generated from a REST API.");
        apiDoc.Save(apiDocPath);

        // -----------------------------------------------------------------
        // 3. Load both documents.
        // -----------------------------------------------------------------
        Document srcDoc = new Document(apiDocPath);
        Document dstDoc = new Document(existingDocPath);

        // -----------------------------------------------------------------
        // 4. Append the API‑generated document to the existing one.
        // -----------------------------------------------------------------
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
        dstDoc.Save(mergedDocPath);

        // -----------------------------------------------------------------
        // 5. Convert the merged document to PDF and encrypt it with a password.
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails("UserPassword123", "OwnerPassword123")
        };
        dstDoc.Save(finalPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 6. Simple validation to ensure the PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(finalPdfPath))
        {
            throw new InvalidOperationException("Failed to create the encrypted PDF.");
        }
    }
}
