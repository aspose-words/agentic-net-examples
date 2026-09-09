using System;
using System.IO;
using System.Net.Http;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputFolder = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputFolder);

        // -----------------------------------------------------------------
        // 1. Create a local DOCX that will act as the existing document.
        // -----------------------------------------------------------------
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the existing document.");

        string existingDocPath = Path.Combine(outputFolder, "Existing.docx");
        existingDoc.Save(existingDocPath);

        // -----------------------------------------------------------------
        // 2. Obtain a DOCX from a web service (download a sample file).
        // -----------------------------------------------------------------
        // Sample DOCX URL – any publicly reachable DOCX file can be used.
        const string sampleDocUrl = "https://filesamples.com/samples/document/docx/sample3.docx";

        using (HttpClient httpClient = new HttpClient())
        {
            HttpResponseMessage response = httpClient.GetAsync(sampleDocUrl).Result;
            response.EnsureSuccessStatusCode();

            byte[] docBytes = response.Content.ReadAsByteArrayAsync().Result;

            using (MemoryStream webDocStream = new MemoryStream(docBytes))
            {
                // Load the downloaded document from the memory stream.
                Document webDoc = new Document(webDocStream);

                // -----------------------------------------------------------------
                // 3. Append the web‑generated document to the existing one.
                // -----------------------------------------------------------------
                existingDoc.AppendDocument(webDoc, ImportFormatMode.KeepSourceFormatting);
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the merged document as DOCX (optional, for verification).
        // -----------------------------------------------------------------
        string mergedDocxPath = Path.Combine(outputFolder, "Merged.docx");
        existingDoc.Save(mergedDocxPath);

        // -----------------------------------------------------------------
        // 5. Convert the merged document to PDF and encrypt it with a password.
        // -----------------------------------------------------------------
        string encryptedPdfPath = Path.Combine(outputFolder, "MergedEncrypted.pdf");

        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // UserPassword is required to open the PDF; OwnerPassword controls permissions.
            EncryptionDetails = new PdfEncryptionDetails("UserPassword123", "OwnerPassword123")
        };

        existingDoc.Save(encryptedPdfPath, pdfOptions);

        // -----------------------------------------------------------------
        // 6. Validation – ensure the encrypted PDF file exists and contains expected text.
        // -----------------------------------------------------------------
        if (!File.Exists(encryptedPdfPath))
            throw new InvalidOperationException("Encrypted PDF was not created.");

        // Load the encrypted PDF using the same password to verify its content.
        LoadOptions loadOptions = new LoadOptions("UserPassword123");
        Document pdfDoc = new Document(encryptedPdfPath, loadOptions);
        string pdfText = pdfDoc.GetText();

        if (!pdfText.Contains("This is the existing document."))
            throw new InvalidOperationException("Merged content is missing in the encrypted PDF.");

        // If execution reaches this point, the process succeeded.
        Console.WriteLine("Document appended and PDF encrypted successfully.");
    }
}
