using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        const string password = "Secret123";
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";

        // Create a simple document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a password‑protected PDF.");

        // Save the document as a PDF with a user password.
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(password, string.Empty)
        };
        doc.Save(pdfPath, pdfSaveOptions);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the protected PDF.");

        // Load the protected PDF using PdfLoadOptions with the password.
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = password
        };
        Document loadedPdf = new Document(pdfPath, loadOptions);

        // Convert the first page of the PDF to PNG.
        loadedPdf.Save(pngPath, SaveFormat.Png);

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("PNG conversion failed.");
    }
}
