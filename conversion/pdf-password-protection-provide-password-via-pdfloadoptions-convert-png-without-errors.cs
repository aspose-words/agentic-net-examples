using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToPngConverter
{
    static void Main()
    {
        // Path to the encrypted PDF file.
        string pdfPath = @"C:\Docs\EncryptedDocument.pdf";

        // Password required to open the PDF.
        string pdfPassword = "myPassword";

        // Folder where the PNG images will be saved.
        string outputFolder = @"C:\Docs\Images\";

        // Ensure the output folder exists.
        System.IO.Directory.CreateDirectory(outputFolder);

        // Configure load options with the password.
        PdfLoadOptions loadOptions = new PdfLoadOptions();
        loadOptions.Password = pdfPassword;

        // Load the PDF document using the load options.
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Convert each page of the PDF to a separate PNG image.
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Set up image save options for PNG format.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Export only the current page using PageSet.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string pngPath = System.IO.Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the page as a PNG image.
            pdfDocument.Save(pngPath, pngOptions);
        }

        Console.WriteLine("PDF conversion to PNG completed successfully.");
    }
}
