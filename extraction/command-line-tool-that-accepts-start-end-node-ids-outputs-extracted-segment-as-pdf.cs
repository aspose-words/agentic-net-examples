using System;
using Aspose.Words;
using Aspose.Words.Saving;

namespace ExtractSegmentTool
{
    class Program
    {
        // Usage:
        // ExtractSegmentTool.exe <inputDocPath> <startPage> <endPage> <outputPdfPath>
        static void Main(string[] args)
        {
            // Validate arguments.
            if (args.Length != 4)
            {
                Console.WriteLine("Incorrect number of arguments.");
                Console.WriteLine("Usage: ExtractSegmentTool.exe <inputDocPath> <startPage> <endPage> <outputPdfPath>");
                return;
            }

            string inputPath = args[0];
            string startPageStr = args[1];
            string endPageStr = args[2];
            string outputPath = args[3];

            if (!int.TryParse(startPageStr, out int startPage) || startPage < 1)
            {
                Console.WriteLine("Start page must be a positive integer.");
                return;
            }

            if (!int.TryParse(endPageStr, out int endPage) || endPage < startPage)
            {
                Console.WriteLine("End page must be an integer greater than or equal to start page.");
                return;
            }

            try
            {
                // Load the source document.
                Document sourceDoc = new Document(inputPath);

                // Convert page numbers (1‑based) to zero‑based index and count.
                int zeroBasedStartIndex = startPage - 1;
                int pageCount = endPage - startPage + 1;

                // Extract the required page range.
                Document extractedDoc = sourceDoc.ExtractPages(zeroBasedStartIndex, pageCount);

                // Save the extracted segment as PDF.
                extractedDoc.Save(outputPath, SaveFormat.Pdf);
                
                Console.WriteLine($"Successfully extracted pages {startPage}-{endPage} to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
