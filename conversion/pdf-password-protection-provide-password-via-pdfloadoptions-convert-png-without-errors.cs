using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToPngConverter
{
    static void Main()
    {
        // Temporary paths.
        string tempDir = Path.Combine(Path.GetTempPath(), "PdfToPngDemo");
        Directory.CreateDirectory(tempDir);

        string pdfPath = Path.Combine(tempDir, "EncryptedDocument.pdf");
        string outputFolder = Path.Combine(tempDir, "Images");
        Directory.CreateDirectory(outputFolder);

        // Password for the encrypted PDF.
        const string pdfPassword = "myPassword";

        // -----------------------------------------------------------------
        // Create a simple document and save it as an encrypted PDF.
        // -----------------------------------------------------------------
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.Writeln("This is a sample PDF page 1.");
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("This is a sample PDF page 2.");

        PdfSaveOptions saveOptions = new PdfSaveOptions
        {
            EncryptionDetails = new PdfEncryptionDetails(pdfPassword, "")
        };
        doc.Save(pdfPath, saveOptions);

        // -----------------------------------------------------------------
        // Load the encrypted PDF using the password.
        // -----------------------------------------------------------------
        PdfLoadOptions loadOptions = new PdfLoadOptions { Password = pdfPassword };
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // -----------------------------------------------------------------
        // Convert each page to a PNG image.
        // -----------------------------------------------------------------
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(pageIndex)
            };

            string pngPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");
            pdfDocument.Save(pngPath, pngOptions);
        }

        Console.WriteLine($"PDF conversion to PNG completed successfully. Files saved in: {outputFolder}");
    }
}
