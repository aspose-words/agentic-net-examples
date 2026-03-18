using System;
using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

namespace BatchHtmlToEpub
{
    public class Converter
    {
        /// <summary>
        /// Converts all .html files in the specified input directory to .epub files in the output directory.
        /// </summary>
        /// <param name="inputFolder">Folder containing source HTML files.</param>
        /// <param name="outputFolder">Folder where the generated EPUB files will be saved.</param>
        public static void BatchConvertHtmlToEpub(string inputFolder, string outputFolder)
        {
            // Ensure the input directory exists.
            if (!Directory.Exists(inputFolder))
                throw new DirectoryNotFoundException($"Input folder not found: {inputFolder}");

            // Create the output directory if it does not exist.
            Directory.CreateDirectory(outputFolder);

            // Get all HTML files in the input folder (non‑recursive).
            string[] htmlFiles = Directory.GetFiles(inputFolder, "*.html");

            foreach (string htmlPath in htmlFiles)
            {
                // Load the HTML document.
                Document doc = new Document(htmlPath);

                // Configure save options for EPUB output.
                HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Epub)
                {
                    // Use UTF‑8 encoding for the EPUB content.
                    Encoding = Encoding.UTF8,

                    // Split the EPUB into separate parts at heading paragraphs (optional, improves readability on some devices).
                    DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph,

                    // Export built‑in and custom document properties into the EPUB.
                    ExportDocumentProperties = true
                };

                // Build the output file name with .epub extension.
                string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".epub";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the document as EPUB using the configured options.
                doc.Save(outputPath, saveOptions);
            }
        }

        // Example usage.
        public static void Main()
        {
            string inputDir = @"C:\HtmlSource";
            string outputDir = @"C:\EpubCollection";

            BatchConvertHtmlToEpub(inputDir, outputDir);

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
