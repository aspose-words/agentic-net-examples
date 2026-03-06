using System;
using System.IO;
using Aspose.Words;

namespace DocumentSplitter
{
    public class Splitter
    {
        /// <summary>
        /// Splits a DOCX document into separate files based on the supplied page ranges.
        /// Each range is defined by a zero‑based start page index and the number of pages to extract.
        /// </summary>
        /// <param name="sourcePath">Full path to the source DOCX file.</param>
        /// <param name="outputDir">Directory where the split documents will be saved.</param>
        /// <param name="pageRanges">
        /// An array of integer pairs: [startIndex, count].
        /// Example: new int[][] { new[] {0, 2}, new[] {2, 3} } extracts pages 1‑2 and 3‑5.
        /// </param>
        public static void SplitDocument(string sourcePath, string outputDir, int[][] pageRanges)
        {
            // Ensure the output directory exists.
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the source document using the Document constructor (lifecycle rule).
            Document sourceDoc = new Document(sourcePath);

            // Iterate over each requested range.
            for (int i = 0; i < pageRanges.Length; i++)
            {
                int startIndex = pageRanges[i][0]; // zero‑based page index
                int count = pageRanges[i][1];      // number of pages to extract

                // Extract the specified pages. This uses the Document.ExtractPages method (feature rule).
                Document part = sourceDoc.ExtractPages(startIndex, count);

                // Build a file name that reflects the range (e.g., "Part_1_2.docx" for pages 1‑2).
                string partFileName = Path.Combine(
                    outputDir,
                    $"Part_{startIndex + 1}_{startIndex + count}.docx");

                // Save the extracted part using the Document.Save method (lifecycle rule).
                part.Save(partFileName);
            }
        }

        // Example usage.
        public static void Main()
        {
            string source = @"C:\Docs\SourceDocument.docx";
            string output = @"C:\Docs\SplitParts";

            // Define page ranges: pages 1‑2, 3‑5, and 6‑6.
            int[][] ranges = new int[][]
            {
                new[] {0, 2}, // pages 1‑2
                new[] {2, 3}, // pages 3‑5
                new[] {5, 1}  // page 6
            };

            SplitDocument(source, output, ranges);
        }
    }
}
