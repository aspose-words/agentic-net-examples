using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Paths for the temporary PDF and the resulting PNG.
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";

        // Passwords for the PDF.
        const string userPassword = "UserPass";
        const string ownerPassword = "OwnerPass";

        // -----------------------------------------------------------------
        // 1. Create a simple document and save it as a password‑protected PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This PDF is protected with a password.");

        // Set encryption details (user password required to open the file).
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(userPassword, ownerPassword)
        };

        sourceDoc.Save(pdfPath, saveOptions);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the protected PDF.");

        // -----------------------------------------------------------------
        // 2. Load the protected PDF using PdfLoadOptions with the password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = userPassword
        };

        Document loadedDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Convert the first page of the PDF to PNG.
        // -----------------------------------------------------------------
        loadedDoc.Save(pngPath, SaveFormat.Png);

        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG conversion did not produce an output file.");

        // The example finishes without waiting for user input.
    }
}
