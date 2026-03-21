using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class AppendAndEncryptPdf
{
    static void Main()
    {
        // Create the existing document in memory.
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the existing document.");

        // Simulate downloading a DOCX from a REST API by creating another document in memory.
        Document apiDoc = CreateSampleApiDocument();

        // Append the API document to the end of the existing document.
        existingDoc.AppendDocument(apiDoc, ImportFormatMode.KeepSourceFormatting);

        // Set up PDF encryption details (user password, no owner password, default permissions).
        PdfEncryptionDetails encryption = new PdfEncryptionDetails("SecretPassword", string.Empty);

        // Configure PDF save options to use the encryption details.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };

        // Save the combined document as an encrypted PDF in the current directory.
        string outputPdfPath = Path.Combine(Directory.GetCurrentDirectory(), "CombinedEncrypted.pdf");
        existingDoc.Save(outputPdfPath, pdfOptions);

        Console.WriteLine($"Encrypted PDF saved to: {outputPdfPath}");
    }

    // Helper method that creates a sample document to simulate the API response.
    private static Document CreateSampleApiDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the document retrieved from the API.");
        return doc;
    }
}
