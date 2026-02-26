using System;
using System.IO;
using Aspose.Words;

namespace Example
{
    public static class DocumentSplitter
    {
        /// <summary>
        /// Splits the specified DOCX document into separate DOCX files, one per page.
        /// </summary>
        /// <param name="sourceFilePath">Full path to the source DOCX file.</param>
        /// <param name="outputDirectory">Folder where the split documents will be saved.</param>
        public static void SplitByPages(string sourceFilePath, string outputDirectory)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Load the source document (lifecycle rule: load).
            Document sourceDoc = new Document(sourceFilePath);

            // Iterate over each page in the source document.
            for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
            {
                // Extract a single page as a new Document (lifecycle rule: create via ExtractPages).
                // Page numbers are 1‑based, so add 1 to the zero‑based index.
                Document pageDoc = sourceDoc.ExtractPages(pageIndex + 1, pageIndex + 1);

                // Build the output file name, e.g., "Document_Page_1.docx".
                string outputFilePath = Path.Combine(
                    outputDirectory,
                    $"{Path.GetFileNameWithoutExtension(sourceFilePath)}_Page_{pageIndex + 1}{Path.GetExtension(sourceFilePath)}");

                // Save the extracted page as a separate DOCX file (lifecycle rule: save).
                pageDoc.Save(outputFilePath);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required for console execution.
        /// </summary>
        static void Main(string[] args)
        {
            // Simple argument handling: first argument = source file, second argument = output folder.
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: DocumentSplitter <source-docx-path> <output-directory>");
                return;
            }

            string sourcePath = args[0];
            string outputDir = args[1];

            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                DocumentSplitter.SplitByPages(sourcePath, outputDir);
                Console.WriteLine($"Document split completed. Files saved to: {outputDir}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during splitting: {ex.Message}");
            }
        }
    }
}
