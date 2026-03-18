using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class PdfToPngConverter
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\sample.pdf";

        // Folder where PNG images will be saved.
        string outputFolder = @"C:\Output\Pages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document with default PDF load options.
        LoadOptions loadOptions = new PdfLoadOptions();
        Document pdfDocument = new Document(pdfPath, loadOptions);

        // Iterate through each page of the PDF.
        for (int pageIndex = 0; pageIndex < pdfDocument.PageCount; pageIndex++)
        {
            // Configure image save options for PNG format.
            ImageSaveOptions imageOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name (1‑based page numbering).
            string outputFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the selected page as a PNG image.
            pdfDocument.Save(outputFile, imageOptions);
        }

        Console.WriteLine("Conversion completed successfully.");
    }
}
