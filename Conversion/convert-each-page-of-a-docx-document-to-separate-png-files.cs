using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPageToPng
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source DOCX file.
            string inputFile = @"C:\Docs\SourceDocument.docx";

            // Folder where the PNG files will be saved.
            string outputFolder = @"C:\Docs\PagesAsPng";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the DOCX document.
            Document doc = new Document(inputFile);

            // Create ImageSaveOptions for PNG format.
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);

            // Iterate through each page in the document.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the PageSet to render only the current page (zero‑based index).
                saveOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name for the current page.
                string outputFile = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

                // Save the current page as a PNG image.
                doc.Save(outputFile, saveOptions);
            }

            Console.WriteLine("All pages have been saved as PNG files.");
        }
    }
}
