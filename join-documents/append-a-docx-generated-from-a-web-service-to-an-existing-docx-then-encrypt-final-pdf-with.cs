using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create a sample existing DOCX file.
        string existingDocPath = Path.Combine(outputDir, "Existing.docx");
        Document existingDoc = new Document();
        DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);
        existingBuilder.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath);

        // Download a DOCX from a web service (simulated by a public URL).
        string webDocUrl = "https://filesamples.com/samples/document/docx/sample3.docx";
        byte[] webDocBytes;
        using (HttpClient httpClient = new HttpClient())
        {
            webDocBytes = httpClient.GetByteArrayAsync(webDocUrl).Result;
        }

        // Load the downloaded DOCX into a Document object.
        Document webDoc;
        using (MemoryStream webStream = new MemoryStream(webDocBytes))
        {
            webStream.Position = 0; // Ensure the stream is at the beginning.
            webDoc = new Document(webStream);
        }

        // Load the existing DOCX.
        Document mergedDoc = new Document(existingDocPath);

        // Append the web‑service document to the existing document.
        mergedDoc.AppendDocument(webDoc, ImportFormatMode.KeepSourceFormatting);

        // Encrypt the final PDF with a password.
        string pdfPath = Path.Combine(outputDir, "MergedEncrypted.pdf");
        PdfEncryptionDetails encryption = new PdfEncryptionDetails("UserPassword", "OwnerPassword");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };
        mergedDoc.Save(pdfPath, pdfOptions);

        // Validate that the PDF was created.
        if (!File.Exists(pdfPath))
        {
            throw new InvalidOperationException("The encrypted PDF was not created.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document merged and encrypted PDF saved to: " + pdfPath);
    }
}
