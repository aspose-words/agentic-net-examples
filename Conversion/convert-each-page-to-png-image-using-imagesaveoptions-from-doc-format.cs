using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsConvertPagesToPng
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOC document.
            string inputDocPath = @"C:\Docs\SourceDocument.doc";

            // Directory where the PNG images will be saved.
            string outputDir = @"C:\Docs\PageImages";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDir);

            // Load the DOC document.
            Document doc = new Document(inputDocPath);

            // Create ImageSaveOptions for PNG format.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

            // Iterate through each page in the document.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the PageSet to render only the current page (zero‑based index).
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name for the current page.
                string outputPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.png");

                // Save the current page as a PNG image.
                doc.Save(outputPath, pngOptions);
            }
        }
    }
}
