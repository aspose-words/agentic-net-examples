using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace WordSplitExample
{
    public static class WordSplitter
    {
        /// <summary>
        /// Splits the source Word document into PDF files, each containing the specified number of pages.
        /// </summary>
        /// <param name="sourcePath">Full path to the source .docx (or other supported) file.</param>
        /// <param name="outputFolder">Folder where the resulting PDF chunks will be saved.</param>
        /// <param name="pagesPerChunk">Number of pages per PDF file (default is 50).</param>
        public static void SplitIntoPdfChunks(string sourcePath, string outputFolder, int pagesPerChunk = 50)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Load the original document.
            Document sourceDoc = new Document(sourcePath);

            // Total number of pages in the source document.
            int totalPages = sourceDoc.PageCount;

            // Iterate over the document in steps of pagesPerChunk.
            for (int startIndex = 0; startIndex < totalPages; startIndex += pagesPerChunk)
            {
                // Determine how many pages to extract for this chunk.
                int count = Math.Min(pagesPerChunk, totalPages - startIndex);

                // Extract the required range of pages.
                Document chunkDoc = sourceDoc.ExtractPages(startIndex, count);

                // Build the output file name (e.g., Part_1.pdf, Part_2.pdf, ...).
                int partNumber = (startIndex / pagesPerChunk) + 1;
                string outputPath = Path.Combine(outputFolder, $"Part_{partNumber}.pdf");

                // Save the extracted chunk as PDF.
                chunkDoc.Save(outputPath, SaveFormat.Pdf);
            }
        }
    }

    class Program
    {
        /// <summary>
        /// Entry point required by the C# compiler. Demonstrates usage of WordSplitter.
        /// </summary>
        /// <param name="args">[0] – source .docx path, [1] – output folder (optional).</param>
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: WordSplitExample <source.docx> [outputFolder]");
                return;
            }

            string sourcePath = args[0];
            string outputFolder = args.Length > 1 ? args[1] : Path.Combine(Path.GetDirectoryName(sourcePath) ?? ".", "PdfChunks");

            try
            {
                WordSplitter.SplitIntoPdfChunks(sourcePath, outputFolder);
                Console.WriteLine($"Document split successfully. PDFs saved to '{outputFolder}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
