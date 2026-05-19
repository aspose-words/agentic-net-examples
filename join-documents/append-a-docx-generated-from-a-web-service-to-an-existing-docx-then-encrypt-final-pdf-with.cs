using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create an existing DOCX file.
        string existingDocPath = Path.Combine(outputDir, "Existing.docx");
        Document existingDoc = new Document();
        DocumentBuilder existingBuilder = new DocumentBuilder(existingDoc);
        existingBuilder.Writeln("This is the existing document.");
        existingDoc.Save(existingDocPath);

        // Simulate a DOCX generated from a web service.
        string webDocPath = Path.Combine(outputDir, "WebGenerated.docx");
        Document webDoc = new Document();
        DocumentBuilder webBuilder = new DocumentBuilder(webDoc);
        webBuilder.Writeln("Content from web service.");
        webDoc.Save(webDocPath);

        // Load the documents.
        Document destinationDoc = new Document(existingDocPath);
        Document sourceDoc = new Document(webDocPath);

        // Append the web-generated document to the existing one.
        destinationDoc.AppendDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);

        // Save the merged document (optional, for verification).
        string mergedDocPath = Path.Combine(outputDir, "Merged.docx");
        destinationDoc.Save(mergedDocPath);

        // Encrypt the final PDF with a password.
        string pdfPath = Path.Combine(outputDir, "MergedEncrypted.pdf");
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails("UserPassword", "OwnerPassword")
        };
        destinationDoc.Save(pdfPath, pdfOptions);

        // Validate that the PDF file exists.
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("The encrypted PDF was not created.", pdfPath);
        }

        // Verify that the PDF can be opened with the correct password.
        LoadOptions loadOptions = new LoadOptions("UserPassword");
        Document loadedPdf = new Document(pdfPath, loadOptions);
        string extractedText = loadedPdf.GetText().Trim();

        // Simple validation to ensure content from both source documents is present.
        if (!extractedText.Contains("This is the existing document.") ||
            !extractedText.Contains("Content from web service."))
        {
            throw new InvalidOperationException("The merged PDF does not contain expected content.");
        }

        // Execution completed successfully.
    }
}
