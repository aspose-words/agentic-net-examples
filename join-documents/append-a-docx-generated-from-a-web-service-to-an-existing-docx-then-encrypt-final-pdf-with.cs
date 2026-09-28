using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Create a temporary working folder.
        string tempDir = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(tempDir);

        // Define file paths for the sample documents and the results.
        string existingDocPath = Path.Combine(tempDir, "ExistingDocument.docx");
        string webServiceDocPath = Path.Combine(tempDir, "WebServiceDocument.docx");
        string mergedDocPath = Path.Combine(tempDir, "MergedDocument.docx");
        string encryptedPdfPath = Path.Combine(tempDir, "MergedDocument_Encrypted.pdf");

        // -------------------------------------------------
        // 1. Create the first sample DOCX (the existing one).
        // -------------------------------------------------
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the original document.");
        existingDoc.Save(existingDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 2. Create the second sample DOCX (simulating a web‑service output).
        // -------------------------------------------------
        Document webServiceDoc = new Document();
        DocumentBuilder webBuilder = new DocumentBuilder(webServiceDoc);
        webBuilder.Writeln("This content comes from a simulated web service.");
        webServiceDoc.Save(webServiceDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 3. Load both documents from disk.
        // -------------------------------------------------
        Document sourceDoc = new Document(existingDocPath);
        Document docToAppend = new Document(webServiceDocPath);

        // -------------------------------------------------
        // 4. Append the second document to the first one,
        //    preserving the source formatting.
        // -------------------------------------------------
        sourceDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);
        sourceDoc.Save(mergedDocPath, SaveFormat.Docx);

        // -------------------------------------------------
        // 5. Convert the merged document to an encrypted PDF.
        //    Use the overload of PdfEncryptionDetails that does not require
        //    specifying the encryption algorithm (defaults to RC4_128).
        // -------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(
                userPassword: "UserPassword123",
                ownerPassword: "OwnerPassword123")
        };
        sourceDoc.Save(encryptedPdfPath, pdfOptions);

        // -------------------------------------------------
        // 6. Verify that the output files were created.
        // -------------------------------------------------
        if (!File.Exists(mergedDocPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedDocPath);

        if (!File.Exists(encryptedPdfPath))
            throw new FileNotFoundException("Encrypted PDF was not created.", encryptedPdfPath);
    }
}
