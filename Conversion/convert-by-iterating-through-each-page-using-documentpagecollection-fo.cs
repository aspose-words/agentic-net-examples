using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsPageIteration
{
    class Program
    {
        static void Main()
        {
            // Path to the source DOCX document.
            string inputPath = @"C:\Docs\SourceDocument.docx";

            // Directory where each extracted page will be saved as a separate DOCX file.
            string outputDir = @"C:\Docs\ExtractedPages";

            // Ensure the output directory exists.
            Directory.CreateDirectory(outputDir);

            // Load the document using the provided Document constructor (load rule).
            Document doc = new Document(inputPath);

            // Iterate through each page using the document's page count.
            // Aspose.Words does not expose a DocumentPageCollection directly,
            // but the page count can be used to access pages one by one.
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Extract a single page (zero‑based index, count = 1) using the provided ExtractPages method.
                Document singlePageDoc = doc.ExtractPages(pageIndex, 1);

                // Build the output file name for the current page.
                string outputPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.docx");

                // Save the extracted page using the provided Save method (file name determines format).
                singlePageDoc.Save(outputPath);
            }

            Console.WriteLine("Pages extracted successfully.");
        }
    }
}
