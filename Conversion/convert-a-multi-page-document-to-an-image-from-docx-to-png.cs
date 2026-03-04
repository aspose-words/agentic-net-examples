using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxToPngConverter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputFile = @"C:\Docs\MultiPageDocument.docx";

        // Directory where PNG images will be saved.
        string outputFolder = @"C:\Docs\Images\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the DOCX document.
        Document doc = new Document(inputFile);

        // Prepare image save options for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Iterate through each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render only the current page.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name (e.g., Page_1.png, Page_2.png, ...).
            string outputFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

            // Save the rendered page as a PNG image.
            doc.Save(outputFile, options);
        }
    }
}
