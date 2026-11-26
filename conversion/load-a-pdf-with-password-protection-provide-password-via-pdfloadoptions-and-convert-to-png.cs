using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for temporary files
        string pdfPath = "protected.pdf";
        string pngPath = "output.png";

        // 1. Create a sample document
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF with password protection.");

        // 2. Save the document as a password‑protected PDF
        PdfEncryptionDetails encryptionDetails = new PdfEncryptionDetails(
            ownerPassword: "owner123",
            userPassword: "user123");

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryptionDetails
        };
        doc.Save(pdfPath, pdfSaveOptions);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("The protected PDF file was not created.");

        // 3. Load the protected PDF using the password
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = "user123"
        };
        Document loadedDoc = new Document(pdfPath, loadOptions);

        // 4. Convert the first page of the PDF to PNG
        ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            PageSet = new PageSet(0) // first page (zero‑based)
        };
        loadedDoc.Save(pngPath, imageOptions);

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG file was not created.");

        // Indicate successful completion
        Console.WriteLine("PDF loaded with password and converted to PNG successfully.");
    }
}
