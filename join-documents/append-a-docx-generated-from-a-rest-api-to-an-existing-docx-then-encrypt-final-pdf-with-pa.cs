using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create a local base DOCX document.
        string baseDocPath = Path.Combine(outputDir, "Base.docx");
        Document baseDoc = new Document();
        DocumentBuilder baseBuilder = new DocumentBuilder(baseDoc);
        baseBuilder.Writeln("This is the base document.");
        baseDoc.Save(baseDocPath);

        // 2. Create a second DOCX document that simulates the file obtained from a REST API.
        string remoteDocPath = Path.Combine(outputDir, "Remote.docx");
        Document remoteDoc = new Document();
        DocumentBuilder remoteBuilder = new DocumentBuilder(remoteDoc);
        remoteBuilder.Writeln("This is the remote document (simulated API response).");
        remoteDoc.Save(remoteDocPath);

        // Load the remote document from the file system (demonstrates loading from a stream as well).
        Document remoteDocLoaded;
        using (MemoryStream remoteStream = new MemoryStream(File.ReadAllBytes(remoteDocPath)))
        {
            remoteStream.Position = 0; // Ensure the stream is at the beginning.
            remoteDocLoaded = new Document(remoteStream);
        }

        // 3. Append the remote document to the base document, preserving its formatting.
        baseDoc.AppendDocument(remoteDocLoaded, ImportFormatMode.KeepSourceFormatting);

        // 4. Save the merged document as a DOCX file.
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        baseDoc.Save(mergedDocPath);

        // Validate that the merged DOCX was created.
        if (!File.Exists(mergedDocPath))
            throw new InvalidOperationException("Failed to create the merged DOCX file.");

        // 5. Convert the merged document to PDF and encrypt it with a password.
        string pdfPath = Path.Combine(outputDir, "MergedEncrypted.pdf");
        PdfEncryptionDetails encryption = new PdfEncryptionDetails("UserPassword", "OwnerPassword");
        PdfSaveOptions pdfOptions = new PdfSaveOptions { EncryptionDetails = encryption };
        baseDoc.Save(pdfPath, pdfOptions);

        // 6. Validate that the encrypted PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the encrypted PDF file.");

        Console.WriteLine("Document merging and PDF encryption completed successfully.");
    }
}
