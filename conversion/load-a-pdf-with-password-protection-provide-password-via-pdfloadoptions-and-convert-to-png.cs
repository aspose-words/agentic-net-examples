using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";
        const string userPassword = "user123";
        const string ownerPassword = "owner123";

        // Create a sample document and save it as a password‑protected PDF.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This PDF is protected with a password.");

        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(ownerPassword, userPassword)
        };
        sourceDoc.Save(pdfPath, pdfSaveOptions);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the protected PDF.");

        // Load the protected PDF using the password.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = userPassword
        };
        Document loadedDoc = new Document(pdfPath, loadOptions);

        // Convert the loaded PDF to PNG.
        ImageSaveOptions pngSaveOptions = new ImageSaveOptions(SaveFormat.Png);
        loadedDoc.Save(pngPath, pngSaveOptions);

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG conversion failed; output file was not created.");

        // Optional cleanup.
        // File.Delete(pdfPath);
        // File.Delete(pngPath);
    }
}
