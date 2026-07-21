using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // Define file names and password.
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";
        const string password = "Secret123";

        // -----------------------------------------------------------------
        // Step 1: Create a simple Word document and save it as a password‑protected PDF.
        // -----------------------------------------------------------------
        Document docToProtect = new Document();
        DocumentBuilder builder = new DocumentBuilder(docToProtect);
        builder.Writeln("This PDF is protected with a password.");

        // Set encryption details: user password required to open the PDF.
        PdfEncryptionDetails encryption = new PdfEncryptionDetails(password, string.Empty);
        PdfSaveOptions saveOptions = new PdfSaveOptions { EncryptionDetails = encryption };
        docToProtect.Save(pdfPath, saveOptions);

        // Verify that the protected PDF was created.
        if (!File.Exists(pdfPath) || new FileInfo(pdfPath).Length == 0)
            throw new InvalidOperationException("Failed to create the password‑protected PDF.");

        // -----------------------------------------------------------------
        // Step 2: Load the protected PDF using PdfLoadOptions with the password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions { Password = password };
        Document loadedDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Step 3: Convert the loaded PDF to PNG (first page only).
        // -----------------------------------------------------------------
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        loadedDoc.Save(pngPath, pngOptions);

        // -----------------------------------------------------------------
        // Validation: ensure the PNG file exists and contains data.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath))
            throw new InvalidOperationException("The PNG output file was not created.");

        if (new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("The PNG output file is empty.");

        // The example completes without requiring any user interaction.
    }
}
