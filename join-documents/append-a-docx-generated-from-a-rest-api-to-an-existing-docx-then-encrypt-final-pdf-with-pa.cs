using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Ensure output directory exists
        string outputDir = "output";
        Directory.CreateDirectory(outputDir);

        // Paths for documents
        string existingDocPath = Path.Combine(outputDir, "existing.docx");
        string mergedDocPath = Path.Combine(outputDir, "merged.docx");
        string finalPdfPath = Path.Combine(outputDir, "final.pdf");

        // 1. Create an existing DOCX file locally
        Document existingDoc = new Document();
        DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);
        existingBuilder.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath);

        // 2. Simulate a REST API that returns a DOCX as a byte array
        byte[] apiDocBytes = GetDocumentFromApi();

        // 3. Load the DOCX returned from the API
        Document apiDoc;
        using (MemoryStream apiStream = new MemoryStream(apiDocBytes))
        {
            apiDoc = new Document(apiStream);
        }

        // 4. Append the API-generated document to the existing document
        existingDoc.AppendDocument(apiDoc, ImportFormatMode.KeepSourceFormatting);
        existingDoc.Save(mergedDocPath);

        // 5. Convert the merged document to PDF and encrypt it with a password
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(
                userPassword: "userPassword",
                ownerPassword: "ownerPassword")
        };
        existingDoc.Save(finalPdfPath, pdfOptions);

        // 6. Validate that the files were created
        if (!File.Exists(existingDocPath))
            throw new FileNotFoundException("Existing DOCX was not created.", existingDocPath);
        if (!File.Exists(mergedDocPath))
            throw new FileNotFoundException("Merged DOCX was not created.", mergedDocPath);
        if (!File.Exists(finalPdfPath))
            throw new FileNotFoundException("Final encrypted PDF was not created.", finalPdfPath);
    }

    // Simulated REST API call that returns a DOCX file as a byte array
    private static byte[] GetDocumentFromApi()
    {
        Document apiDoc = new Document();
        DocumentBuilder apiBuilder = new DocumentBuilder(apiDoc);
        apiBuilder.Writeln("This is the document generated from the REST API.");
        using (MemoryStream stream = new MemoryStream())
        {
            apiDoc.Save(stream, SaveFormat.Docx);
            return stream.ToArray();
        }
    }
}
