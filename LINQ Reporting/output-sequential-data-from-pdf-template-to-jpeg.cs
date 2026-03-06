using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class PdfToJpegConverter
{
    static void Main()
    {
        // Path to the PDF template.
        string pdfPath = @"Template.pdf";

        // Directory where JPEG images will be saved.
        string outputFolder = @"Output";
        Directory.CreateDirectory(outputFolder);

        // Load the PDF document.
        Document doc = new Document(pdfPath);

        // Configure image save options for JPEG output.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);
        jpegOptions.Resolution = 300;      // Set desired DPI.
        jpegOptions.JpegQuality = 90;      // Set JPEG quality (0‑100).

        // Render each page of the PDF to a separate JPEG file.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Select the current page (zero‑based index).
            jpegOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name.
            string jpegPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.jpg");

            // Save the selected page as JPEG.
            doc.Save(jpegPath, jpegOptions);
        }
    }
}
