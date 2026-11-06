using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare a temporary working directory.
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinDemo");
        Directory.CreateDirectory(workDir);

        // Paths for the sample existing DOCX and the final encrypted PDF.
        string existingDocPath = Path.Combine(workDir, "Existing.docx");
        string outputPdfPath = Path.Combine(workDir, "CombinedEncrypted.pdf");

        // -----------------------------------------------------------------
        // 1. Create a sample existing DOCX document.
        // -----------------------------------------------------------------
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the content of the existing DOCX document.");
        existingDoc.Save(existingDocPath, SaveFormat.Docx);

        // -----------------------------------------------------------------
        // 2. Simulate a web service that returns a DOCX document.
        //    Here we create the document in memory and obtain a stream.
        // -----------------------------------------------------------------
        using (MemoryStream webServiceStream = new MemoryStream())
        {
            Document webDoc = new Document();
            DocumentBuilder webBuilder = new DocumentBuilder(webDoc);
            webBuilder.Writeln("This content comes from the simulated web service DOCX.");
            webDoc.Save(webServiceStream, SaveFormat.Docx);
            webServiceStream.Position = 0; // Reset before loading.

            // Load the document returned by the web service.
            Document webDocFromStream = new Document(webServiceStream);

            // -----------------------------------------------------------------
            // 3. Append the web service document to the existing document.
            // -----------------------------------------------------------------
            existingDoc.AppendDocument(webDocFromStream, ImportFormatMode.KeepSourceFormatting);
        }

        // -----------------------------------------------------------------
        // 4. Save the combined document as an encrypted PDF.
        // -----------------------------------------------------------------
        PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails(
            userPassword: "UserPassword123",
            ownerPassword: "OwnerPassword123");

        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryptionDetails
        };

        existingDoc.Save(outputPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 5. Validate that the encrypted PDF was created.
        // -----------------------------------------------------------------
        if (!File.Exists(outputPdfPath))
        {
            throw new InvalidOperationException($"Failed to create the encrypted PDF at '{outputPdfPath}'.");
        }

        Console.WriteLine($"Encrypted PDF created successfully at: {outputPdfPath}");
    }
}
