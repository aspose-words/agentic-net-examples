using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class SplitDocumentByRanges
{
    static void Main()
    {
        // Create a sample document with 7 pages in the temp folder.
        string tempFolder = Path.GetTempPath();
        string sourcePath = Path.Combine(tempFolder, "SourceDocument.docx");

        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 7; i++)
        {
            builder.Writeln($"This is page {i}.");
            if (i < 7)
                builder.InsertBreak(BreakType.PageBreak);
        }

        doc.Save(sourcePath);

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

            // Create a PageRange object.
            PageRange pageRange = new PageRange(startPage, endPage);

            // Create a PageSet that contains this single range.
            PageSet pageSet = new PageSet(pageRange);

            // Configure PDF save options and assign the PageSet.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PageSet = pageSet
            };

            // Build the output file name in the temp folder.
            string outputPath = Path.Combine(tempFolder, $"Output_Range_{i + 1}.pdf");

            // Save the selected pages as a PDF.
            doc.Save(outputPath, pdfOptions);

            Console.WriteLine($"Saved pages {bounds[0]}-{bounds[1]} to {outputPath}");
        }
    }
}
