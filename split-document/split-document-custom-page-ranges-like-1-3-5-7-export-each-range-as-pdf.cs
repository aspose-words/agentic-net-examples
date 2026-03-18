using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByRanges
{
    static void Main()
    {
        // Path to the source document.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the document (create/load rule).
        Document doc = new Document(sourcePath);

        // Define the custom page ranges (1‑based, inclusive).
        // Example: "1-3,5-7" means pages 1‑3 and 5‑7.
        string ranges = "1-3,5-7";

        // Split the string into individual range specifications.
        string[] rangeParts = ranges.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

        // Process each range separately.
        for (int i = 0; i < rangeParts.Length; i++)
        {
            // Parse start and end page numbers (convert to zero‑based indices).
            string[] bounds = rangeParts[i].Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (bounds.Length != 2)
                continue; // Invalid range – skip.

            int startPage = int.Parse(bounds[0].Trim()) - 1; // zero‑based
            int endPage   = int.Parse(bounds[1].Trim()) - 1; // zero‑based

            // Create a PageRange object (constructor rule).
            PageRange pageRange = new PageRange(startPage, endPage);

            // Create a PageSet that contains this single range (constructor rule).
            PageSet pageSet = new PageSet(pageRange);

            // Configure PDF save options and assign the PageSet (property rule).
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PageSet = pageSet
            };

            // Build the output file name.
            string outputPath = $@"C:\Docs\Output_Range_{i + 1}.pdf";

            // Save the selected pages as a PDF (save rule).
            doc.Save(outputPath, pdfOptions);
        }
    }
}
