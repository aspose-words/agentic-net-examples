using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Loading;

public class Program
{
    public static void Main()
    {
        // Define file names in the current directory.
        string baseDir = Directory.GetCurrentDirectory();
        string existingDocPath = Path.Combine(baseDir, "Existing.docx");
        string generatedDocPath = Path.Combine(baseDir, "Generated.docx");
        string mergedPdfPath = Path.Combine(baseDir, "MergedEncrypted.pdf");

        // -------------------------------------------------
        // 1. Create an existing DOCX file.
        // -------------------------------------------------
        Document existingDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(existingDoc);
        builder.Writeln("This is the original document.");
        existingDoc.Save(existingDocPath);

        // -------------------------------------------------
        // 2. Simulate a REST API that returns a DOCX.
        //    Here we simply create another document.
        // -------------------------------------------------
        Document generatedDoc = new Document();
        DocumentBuilder genBuilder = new DocumentBuilder(generatedDoc);
        genBuilder.Writeln("Content generated from a REST API.");
        generatedDoc.Save(generatedDocPath);

        // -------------------------------------------------
        // 3. Load both documents.
        // -------------------------------------------------
        Document srcDoc = new Document(generatedDocPath);
        Document dstDoc = new Document(existingDocPath);

        // -------------------------------------------------
        // 4. Append the generated document to the existing one.
        // -------------------------------------------------
        dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

        // -------------------------------------------------
        // 5. Save the merged document as an encrypted PDF.
        // -------------------------------------------------
        // Define user and owner passwords.
        const string userPassword = "UserPass123";
        const string ownerPassword = "OwnerPass123";

        // Set up PDF encryption details.
        PdfEncryptionDetails encryption = new PdfEncryptionDetails(userPassword, ownerPassword);

        // Create PDF save options with the encryption details.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };

        // Save the merged document to PDF with encryption.
        dstDoc.Save(mergedPdfPath, pdfOptions);

        // -------------------------------------------------
        // 6. Simple validation that the output file exists.
        // -------------------------------------------------
        if (!File.Exists(mergedPdfPath))
        {
            throw new FileNotFoundException("The encrypted PDF was not created.", mergedPdfPath);
        }

        // The program finishes without requiring any user interaction.
    }
}
