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
        const string pdfPath = "protected.pdf";
        const string pngPath = "output.png";

        // -----------------------------------------------------------------
        // 1. Create a simple Word document and save it as a password‑protected PDF.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("This PDF is protected with a password.");

        // Set PDF encryption (user password required to open the file)
        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails("SecretPassword", string.Empty)
        };
        sourceDoc.Save(pdfPath, saveOptions);

        if (!File.Exists(pdfPath))
            throw new InvalidOperationException("Failed to create the protected PDF.");

        // -----------------------------------------------------------------
        // 2. Load the protected PDF using PdfLoadOptions with the correct password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions
        {
            Password = "SecretPassword"
        };
        Document protectedDoc = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // 3. Convert the loaded PDF to a PNG image.
        // -----------------------------------------------------------------
        protectedDoc.Save(pngPath, SaveFormat.Png);

        // -----------------------------------------------------------------
        // 4. Validate that the PNG file was created.
        // -----------------------------------------------------------------
        if (!File.Exists(pngPath) || new FileInfo(pngPath).Length == 0)
            throw new InvalidOperationException("PNG conversion failed; output file is missing or empty.");

        // Example completed successfully.
        Console.WriteLine("PDF loaded with password and converted to PNG successfully.");
    }
}
