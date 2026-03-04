using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocToPngConverter
{
    /// <summary>
    /// Converts each page of a DOC/DOCX document to a separate PNG image.
    /// </summary>
    /// <param name="inputFilePath">Full path to the source Word document.</param>
    /// <param name="outputFolderPath">Folder where PNG files will be written.</param>
    public static void ConvertPagesToPng(string inputFilePath, string outputFolderPath)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolderPath);

        // Load the Word document from the specified file.
        Document doc = new Document(inputFilePath);

        // Create ImageSaveOptions for PNG output.
        ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);
        // Optional: set resolution (dpi) for higher quality images.
        pngOptions.Resolution = 300;

        // Iterate through all pages in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Configure the options to render only the current page.
            pngOptions.PageSet = new PageSet(pageIndex);

            // Build the output file name (Page_1.png, Page_2.png, ...).
            string outputFilePath = Path.Combine(outputFolderPath, $"Page_{pageIndex + 1}.png");

            // Save the current page as a PNG image.
            doc.Save(outputFilePath, pngOptions);
        }
    }

    // Example usage.
    static void Main()
    {
        string sourceDoc = @"C:\Docs\Sample.docx";
        string pngFolder = @"C:\Docs\PagesAsPng";

        ConvertPagesToPng(sourceDoc, pngFolder);
    }
}
