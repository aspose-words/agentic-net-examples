using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToPngSequence
{
    static void Main()
    {
        // Path to the source PDF file.
        string pdfPath = @"C:\Input\source.pdf";

        // Folder where the PNG images will be saved.
        string outputFolder = @"C:\Output\Images";

        // Ensure the output folder exists.
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document pdfDoc = new Document(pdfPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png)
        {
            // Set the background (paper) color to transparent.
            PaperColor = Color.Transparent,

            // Optional: set resolution if needed (default is 96 DPI).
            // Resolution = 300
        };

        // Iterate through each page of the PDF.
        for (int pageIndex = 0; pageIndex < pdfDoc.PageCount; pageIndex++)
        {
            // Configure the options to render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image with transparent background.
            pdfDoc.Save(outputPath, pngOptions);
        }

        Console.WriteLine("PDF pages have been saved as PNG images with transparent background.");
    }
}
