using System;
using System.Collections.Generic;
using Aspose.Words;

namespace DocumentSplitter
{
    class Program
    {
        static void Main()
        {
            // Load the source DOCX document.
            // This uses the Document(string) constructor – a provided lifecycle rule.
            Document sourceDoc = new Document("InputDocument.docx");

            // Define the page ranges to split the document into.
            // Each tuple contains the start page (1‑based) and end page (1‑based) inclusive.
            List<(int startPage, int endPage)> pageRanges = new List<(int, int)>
            {
                (1, 2),   // Pages 1‑2
                (3, 5),   // Pages 3‑5
                (6, 6)    // Page 6 only
            };

            // Iterate over the defined ranges and extract each part.
            for (int i = 0; i < pageRanges.Count; i++)
            {
                var range = pageRanges[i];

                // Convert to zero‑based index required by ExtractPages.
                int zeroBasedStart = range.startPage - 1;
                int pageCount = range.endPage - range.startPage + 1;

                // Extract the specified pages.
                // This uses the Document.ExtractPages(int, int) method – a provided rule.
                Document partDoc = sourceDoc.ExtractPages(zeroBasedStart, pageCount);

                // Save the extracted part to a new DOCX file.
                // This uses the Document.Save(string) method – a provided rule.
                string outputPath = $"Output_Part_{i + 1}_Pages_{range.startPage}_to_{range.endPage}.docx";
                partDoc.Save(outputPath);
            }

            Console.WriteLine("Document split completed.");
        }
    }
}
