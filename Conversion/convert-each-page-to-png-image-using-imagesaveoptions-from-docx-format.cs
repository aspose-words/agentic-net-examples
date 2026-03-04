using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace Example
{
    class DocxToPngConverter
    {
        /// <summary>
        /// Converts each page of a DOCX document to a separate PNG image.
        /// </summary>
        /// <param name="inputDocxPath">Full path to the source DOCX file.</param>
        /// <param name="outputFolder">Folder where PNG files will be written.</param>
        public void Convert(string inputDocxPath, string outputFolder)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolder);

            // Load the DOCX document.
            Document doc = new Document(inputDocxPath);

            // Prepare image save options for PNG format.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

            // Iterate through all pages in the document.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Render only the current page (zero‑based index).
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name for the current page.
                string outputPath = Path.Combine(outputFolder, $"Page_{pageIndex + 1}.png");

                // Save the single page as a PNG image.
                doc.Save(outputPath, pngOptions);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for a console application.
        /// </summary>
        static void Main(string[] args)
        {
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: Example <input-docx-path> <output-folder>");
                return;
            }

            string inputPath = args[0];
            string outputFolder = args[1];

            var converter = new DocxToPngConverter();
            converter.Convert(inputPath, outputFolder);

            Console.WriteLine("Conversion completed.");
        }
    }
}
