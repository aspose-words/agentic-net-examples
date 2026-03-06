using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SupportedImageFormatsForDoc
{
    static void Main()
    {
        // Path to the source DOC document.
        string inputPath = @"C:\Docs\SourceDocument.doc";

        // Load the DOC document.
        Document doc = new Document(inputPath);

        // Directory where the converted images will be saved.
        string outputDir = @"C:\Docs\ConvertedImages\";

        // Ensure the output directory exists.
        Directory.CreateDirectory(outputDir);

        // List of image SaveFormat values supported for rendering.
        SaveFormat[] imageFormats = new SaveFormat[]
        {
            SaveFormat.Png,
            SaveFormat.Jpeg,
            SaveFormat.Bmp,
            SaveFormat.Gif,
            SaveFormat.Tiff,
            SaveFormat.Emf,
            SaveFormat.Eps,
            SaveFormat.WebP,
            SaveFormat.Svg
        };

        // Iterate over each page of the document.
        for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
        {
            // Render the current page to each supported image format.
            foreach (SaveFormat imgFormat in imageFormats)
            {
                // Configure ImageSaveOptions for the desired format.
                ImageSaveOptions options = new ImageSaveOptions(imgFormat)
                {
                    // Render only the current page.
                    PageSet = new PageSet(pageIndex)
                };

                // Build the output file name: Document_Page{n}_{format}.ext
                string extension = FileFormatUtil.SaveFormatToExtension(imgFormat);
                string outputPath = Path.Combine(
                    outputDir,
                    $"Document_Page{pageIndex + 1}_{imgFormat}{extension}");

                // Save the page as an image.
                doc.Save(outputPath, options);
            }
        }

        Console.WriteLine("Conversion completed.");
    }
}
