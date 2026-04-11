using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define passwords.
        const string userPassword = "Secret123";
        const string ownerPassword = "";

        // Prepare file paths.
        string workDir = Directory.GetCurrentDirectory();
        string encryptedPdfPath = Path.Combine(workDir, "EncryptedDocument.pdf");
        string pngOutputPath = Path.Combine(workDir, "ConvertedPage.png");

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as a password‑protected PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This PDF is protected with a password.");

        // Set up PDF encryption (user password required to open the file).
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(userPassword, ownerPassword)
        };
        doc.Save(encryptedPdfPath, pdfSaveOptions);

        // Verify that the encrypted PDF was created.
        if (!File.Exists(encryptedPdfPath))
            throw new FileNotFoundException("Failed to create the encrypted PDF.", encryptedPdfPath);

        // -----------------------------------------------------------------
        // 2. Load the encrypted PDF using PdfLoadOptions with the correct password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = userPassword
        };
        Document loadedDoc = new Document(encryptedPdfPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Convert the first page of the PDF to PNG.
        // -----------------------------------------------------------------
        loadedDoc.Save(pngOutputPath, SaveFormat.Png);

        // Verify that the PNG file was created.
        if (!File.Exists(pngOutputPath))
            throw new FileNotFoundException("PNG conversion failed.", pngOutputPath);

        // Indicate successful completion.
        Console.WriteLine("PDF loaded with password and converted to PNG successfully.");
    }
}
