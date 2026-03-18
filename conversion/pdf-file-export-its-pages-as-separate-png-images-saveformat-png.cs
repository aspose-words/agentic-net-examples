using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = "input.pdf";

        // Folder where the PNG images will be saved.
        string outputFolder = "Pages";
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document. PdfLoadOptions is used because the source format is PDF.
        Document doc = new Document(pdfPath, new PdfLoadOptions());

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Configure image save options to render PNG images.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                // Render only the current page.
                PageSet = new PageSet(pageIndex)
            };

            // Build the output file name (Page_1.png, Page_2.png, ...).
            string outFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image.
            doc.Save(outFile, pngOptions);
        }
    }
}
