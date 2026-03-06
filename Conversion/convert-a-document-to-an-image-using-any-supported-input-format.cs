using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentToImageConverter
{
    // Converts each page of a Word document to separate PNG images.
    // inputPath  - path to the source document (any format supported by Aspose.Words).
    // outputDir  - directory where the PNG files will be written.
    public static void ConvertToPngImages(string inputPath, string outputDir)
    {
        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // Load the document from the specified file.
        Document doc = new Document(inputPath);

        // Create ImageSaveOptions for PNG format.
        ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

        // Iterate through all pages in the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render only the current page.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name, e.g., "Document_Page1.png".
            string outputPath = Path.Combine(outputDir,
                $"{Path.GetFileNameWithoutExtension(inputPath)}_Page{pageIndex + 1}.png");

            // Save the current page as an image using the Save method that accepts SaveOptions.
            doc.Save(outputPath, options);
        }
    }

    // Example usage.
    static void Main()
    {
        string sourceFile = @"C:\Docs\SampleDocument.docx";   // any supported format
        string imagesFolder = @"C:\Docs\ConvertedImages";

        ConvertToPngImages(sourceFile, imagesFolder);

        Console.WriteLine("Conversion completed.");
    }
}
