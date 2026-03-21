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

                // Build the output file name with .epub extension.
                string outputFileName = Path.GetFileNameWithoutExtension(htmlPath) + ".epub";
                string outputPath = Path.Combine(outputFolder, outputFileName);

                // Save the document as EPUB.
                doc.Save(outputPath, SaveFormat.Epub);
            }
        }

        // Example usage.
        public static void Main()
        {
            // Use directories relative to the current working directory so they always exist.
            string baseDir = Directory.GetCurrentDirectory();
            string inputDir = Path.Combine(baseDir, "HtmlSource");
            string outputDir = Path.Combine(baseDir, "EpubCollection");

            // Ensure input directory exists.
            Directory.CreateDirectory(inputDir);
            // Ensure output directory exists.
            Directory.CreateDirectory(outputDir);

            // If there are no HTML files, create a simple sample file.
            if (Directory.GetFiles(inputDir, "*.html").Length == 0)
            {
                string sampleHtmlPath = Path.Combine(inputDir, "Sample.html");
                File.WriteAllText(sampleHtmlPath,
                    "<html><head><title>Sample</title></head><body><h1>Hello, EPUB!</h1><p>This is a sample HTML file.</p></body></html>",
                    Encoding.UTF8);
            }

            BatchConvertHtmlToEpub(inputDir, outputDir);

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
