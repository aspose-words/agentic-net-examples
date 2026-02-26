using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExamples
{
    /// <summary>
    /// Converts each page of a DOC/DOCX document to a separate PNG image.
    /// </summary>
    public class DocToPngConverter
    {
        /// <summary>
        /// Converts each page of a DOC/DOCX document to a separate PNG image.
        /// </summary>
        /// <param name="inputFilePath">Full path to the source Word document.</param>
        /// <param name="outputFolderPath">Folder where PNG files will be written. It will be created if it does not exist.</param>
        public void Convert(string inputFilePath, string outputFolderPath)
        {
            // Ensure the output directory exists.
            Directory.CreateDirectory(outputFolderPath);

            // Load the document from the specified file.
            Document doc = new Document(inputFilePath);

            // Prepare ImageSaveOptions for PNG output.
            ImageSaveOptions pngOptions = new ImageSaveOptions(SaveFormat.Png);

            // Iterate through all pages (PageCount is 1‑based).
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Set the PageSet to the current zero‑based page index.
                pngOptions.PageSet = new PageSet(pageIndex);

                // Build the output file name, e.g. "Document_Page_1.png".
                string outFile = Path.Combine(outputFolderPath,
                    $"{Path.GetFileNameWithoutExtension(inputFilePath)}_Page_{pageIndex + 1}.png");

                // Save the single page as a PNG image.
                doc.Save(outFile, pngOptions);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point of the console application.
        /// </summary>
        static void Main(string[] args)
        {
            // Simple argument handling – you can replace these paths with your own.
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: DocToPngConverter <input-doc-path> <output-folder-path>");
                return;
            }

            string inputPath = args[0];
            string outputPath = args[1];

            try
            {
                var converter = new DocToPngConverter();
                converter.Convert(inputPath, outputPath);
                Console.WriteLine($"Conversion completed. PNG files are saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
