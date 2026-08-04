using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    // Password to protect the final PDF.
    private const string PdfPassword = "Secret123";

    // Entry point.
    public static void Main()
    {
        // Folder for all generated files.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // 1. Create the base DOCX document.
        string baseDocPath = Path.Combine(outputDir, "Base.docx");
        Document baseDoc = CreateBaseDocument();
        baseDoc.Save(baseDocPath);

        // 2. Simulate a REST API that returns a DOCX document.
        //    Here we generate the document locally and obtain it as a stream.
        using MemoryStream apiDocStream = GenerateDocumentFromApi();

        // 3. Load the API-generated DOCX from the stream.
        apiDocStream.Position = 0;
        Document apiDoc = new Document(apiDocStream);

        // 4. Append the API document to the base document.
        baseDoc.AppendDocument(apiDoc, ImportFormatMode.KeepSourceFormatting);

        // 5. Save the combined document.
        string combinedDocPath = Path.Combine(outputDir, "Combined.docx");
        baseDoc.Save(combinedDocPath);

        // 6. Convert the combined DOCX to PDF with password protection.
        string encryptedPdfPath = Path.Combine(outputDir, "FinalEncrypted.pdf");
        SavePdfWithPassword(baseDoc, encryptedPdfPath, PdfPassword);

        // 7. Validate that the output files were created.
        ValidateFileExists(baseDocPath);
        ValidateFileExists(combinedDocPath);
        ValidateFileExists(encryptedPdfPath);
    }

    // Creates a simple DOCX document that will serve as the base file.
    private static Document CreateBaseDocument()
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is the base document.");
        builder.Writeln("It will have content appended from the API-generated document.");
        return doc;
    }

    // Simulates a REST API call that returns a DOCX file.
    // The method builds a document, saves it to a MemoryStream, and returns the stream.
    private static MemoryStream GenerateDocumentFromApi()
    {
        Document apiDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(apiDoc);
        builder.Writeln("Document generated from a simulated REST API.");
        builder.Writeln("Additional API content goes here.");

        MemoryStream stream = new MemoryStream();
        apiDoc.Save(stream, SaveFormat.Docx);
        return stream;
    }

    // Saves the provided document as a password‑protected PDF.
    private static void SavePdfWithPassword(Document doc, string outputPath, string password)
    {
        // Configure PDF encryption details.
        PdfEncryptionDetails encryption = new PdfEncryptionDetails(userPassword: password, ownerPassword: string.Empty);
        PdfSaveOptions options = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };

        doc.Save(outputPath, options);
    }

    // Throws an exception if the specified file does not exist.
    private static void ValidateFileExists(string filePath)
    {
        if (!File.Exists(filePath))
        {
            throw new FileNotFoundException($"Expected file was not created: {filePath}");
        }
    }
}
