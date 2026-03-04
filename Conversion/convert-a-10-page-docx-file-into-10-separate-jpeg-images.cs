using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToJpegPages
{
    static void Main()
    {
        // Path to the source DOCX file (must contain 10 pages).
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Directory where the JPEG images will be saved.
        string outputDir = @"C:\Docs\PagesAsJpeg";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the DOCX document.
        Document doc = new Document(sourcePath);

        // Prepare image save options for JPEG format.
        ImageSaveOptions jpegOptions = new ImageSaveOptions(SaveFormat.Jpeg);

        // Loop through each page in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render only the current page.
            jpegOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name: Page1.jpg, Page2.jpg, ...
            string outputPath = Path.Combine(outputDir, $"Page{pageIndex + 1}.jpg");

            // Save the single page as a JPEG image.
            doc.Save(outputPath, jpegOptions);
        }
    }
}
