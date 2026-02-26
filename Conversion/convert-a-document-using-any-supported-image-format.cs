using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentToImageConverter
{
    static void Main()
    {
        // Path to the source document (any format supported by Aspose.Words).
        string sourcePath = @"C:\Docs\SampleDocument.docx";

        // Folder where the resulting images will be saved.
        string outputFolder = @"C:\Docs\ConvertedImages";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputFolder);

        // Load the document from the file system.
        Document doc = new Document(sourcePath);

        // Choose the desired image format (e.g., PNG, JPEG, BMP, etc.).
        // Here we use PNG as an example.
        SaveFormat imageFormat = SaveFormat.Png;

        // Configure image save options.
        ImageSaveOptions options = new ImageSaveOptions(imageFormat)
        {
            // Optional: set resolution, quality, etc.
            Resolution = 300,          // 300 DPI
            JpegQuality = 90          // Ignored for PNG but kept for completeness
        };

        // Iterate through each page of the document and save it as a separate image.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render only the current page.
            options.PageSet = new PageSet(pageIndex);

            // Build the output file name.
            string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.{imageFormat.ToString().ToLower()}");

            // Save the page as an image.
            doc.Save(outputPath, options);
        }

        Console.WriteLine("Document conversion to images completed.");
    }
}
