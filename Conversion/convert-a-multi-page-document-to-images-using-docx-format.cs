using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsImageExport
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\MultiPageDocument.docx";

            // Folder where the rendered images will be saved.
            string outputFolder = @"C:\Docs\Images\";

            // Ensure the output folder exists.
            System.IO.Directory.CreateDirectory(outputFolder);

            // Load the DOCX document.
            Document doc = new Document(inputPath);

            // Create an ImageSaveOptions object for PNG images.
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png)
            {
                // Optional: set resolution (dpi) and image size.
                Resolution = 300,
                ImageSize = new Size(1240, 1754) // A4 at 300 dpi.
            };

            // Render each page of the document to a separate image file.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the PageSet to the current zero‑based page index.
                options.PageSet = new PageSet(pageIndex);

                // Build the output file name, e.g., Page_1.png, Page_2.png, ...
                string outputPath = System.IO.Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

                // Save the current page as an image.
                doc.Save(outputPath, options);
            }

            Console.WriteLine("Document pages have been successfully exported as images.");
        }
    }
}
