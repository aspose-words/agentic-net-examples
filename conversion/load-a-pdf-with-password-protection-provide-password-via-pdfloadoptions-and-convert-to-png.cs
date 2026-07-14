using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // File names and password.
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";
        const string password = "Secret123";

        // -----------------------------------------------------------------
        // Step 1: Create a simple document and save it as a password‑protected PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This PDF is protected with a password.");

        // The owner password must be different from the user password.
        // An empty owner password disables owner‑level restrictions.
        PdfEncryptionDetails encryption = new PdfEncryptionDetails(userPassword: password, ownerPassword: string.Empty);
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EncryptionDetails = encryption
        };

        sourceDoc.Save(pdfPath, saveOptions);

        // Verify that the protected PDF was created.
        if (!File.Exists(pdfPath))
            throw new InvalidOperationException($"Failed to create the PDF file '{pdfPath}'.");

        // -----------------------------------------------------------------
        // Step 2: Load the password‑protected PDF using PdfLoadOptions.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = password
        };

        Document protectedDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 3: Convert the loaded PDF to PNG.
        // -----------------------------------------------------------------
        protectedDoc.Save(pngPath, SaveFormat.Png);

        // Verify that the PNG was created.
        if (!File.Exists(pngPath))
            throw new InvalidOperationException($"Failed to create the PNG file '{pngPath}'.");

        // Indicate success (no interactive I/O required).
        Console.WriteLine("PDF successfully converted to PNG.");
    }
}
