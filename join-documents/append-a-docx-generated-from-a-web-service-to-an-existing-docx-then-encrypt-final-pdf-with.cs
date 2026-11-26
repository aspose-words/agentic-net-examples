using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare temporary working directory
        string workDir = Path.Combine(Path.GetTempPath(), "AsposeJoinExample");
        Directory.CreateDirectory(workDir);

        // Paths for the documents
        string existingDocPath = Path.Combine(workDir, "Existing.docx");
        string webGeneratedDocPath = Path.Combine(workDir, "WebGenerated.docx");
        string mergedDocPath = Path.Combine(workDir, "Merged.docx");
        string encryptedPdfPath = Path.Combine(workDir, "Encrypted.pdf");

        // 1. Create the existing DOCX
        Document existingDoc = new Document();
        DocumentBuilder builder1 = new DocumentBuilder(existingDoc);
        builder1.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath, SaveFormat.Docx);

        // 2. Simulate a DOCX generated from a web service
        Document webGeneratedDoc = new Document();
        DocumentBuilder builder2 = new DocumentBuilder(webGeneratedDoc);
        builder2.Writeln("This is the document generated from a web service.");
        webGeneratedDoc.Save(webGeneratedDocPath, SaveFormat.Docx);

        // 3. Load the documents (optional, they are already in memory)
        Document sourceDoc = new Document(existingDocPath);
        Document docToAppend = new Document(webGeneratedDocPath);

        // 4. Append the web‑generated document to the existing one
        sourceDoc.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);
        sourceDoc.Save(mergedDocPath, SaveFormat.Docx);

        // 5. Convert the merged document to an encrypted PDF
        PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails("user123", "owner123");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryptionDetails
        };

        sourceDoc.Save(encryptedPdfPath, pdfOptions);

        // 6. Validate that the output files were created
        if (!File.Exists(mergedDocPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedDocPath);
        if (!File.Exists(encryptedPdfPath))
            throw new FileNotFoundException("Encrypted PDF was not created.", encryptedPdfPath);

        // Confirmation (no interactive input)
        Console.WriteLine($"Merged document saved to: {mergedDocPath}");
        Console.WriteLine($"Encrypted PDF saved to: {encryptedPdfPath}");
    }
}
