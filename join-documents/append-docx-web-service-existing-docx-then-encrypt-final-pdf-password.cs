using System;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create the first (existing) document in memory.
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the content of the existing document.");

        // Create the second (web‑generated) document in memory.
        Document webDoc = new Document();
        DocumentBuilder webBuilder = new DocumentBuilder(webDoc);
        webBuilder.Writeln("This is the content of the web‑generated document.");

        // Append the web‑generated document to the existing one, preserving source formatting.
        existingDoc.AppendDocument(webDoc, ImportFormatMode.KeepSourceFormatting);

        // Configure PDF encryption: user password, owner password and desired permissions.
        PdfEncryptionDetails encryption = new PdfEncryptionDetails(
            userPassword: "UserPass123",
            ownerPassword: "OwnerPass123",
            permissions: PdfPermissions.ModifyAnnotations | PdfPermissions.Printing);

        // Create PDF save options and assign the encryption details.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };

        // Save the combined document as an encrypted PDF.
        const string outputPdfPath = "CombinedEncrypted.pdf";
        existingDoc.Save(outputPdfPath, pdfOptions);

        Console.WriteLine($"PDF saved to '{outputPdfPath}' with encryption.");
    }
}
