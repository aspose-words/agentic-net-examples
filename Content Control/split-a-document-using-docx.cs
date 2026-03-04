using System;
using System.IO;
using Aspose.Words;

class SplitDocxExample
{
    static void Main()
    {
        // Load the source DOCX document.
        // The Document constructor is the provided creation/loading rule.
        Document sourceDoc = new Document(@"C:\Input\SourceDocument.docx");

        // Determine the total number of pages in the document.
        // This information is needed to split the document page‑by‑page.
        int pageCount = sourceDoc.PageCount;

        // Create an output folder for the split parts.
        string outputFolder = @"C:\Output\SplitParts";
        Directory.CreateDirectory(outputFolder);

        // Loop through each page and extract it as a separate document.
        // The ExtractPages method returns a new Document containing the specified page range.
        for (int pageIndex = 1; pageIndex <= pageCount; pageIndex++)
        {
            // Extract a single page (pageIndex) from the source document.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Build a file name for the extracted part.
            string partFileName = Path.Combine(outputFolder, $"Part_{pageIndex}.docx");

            // Save the extracted page as a separate DOCX file.
            // The Save method is the provided saving rule.
            pageDoc.Save(partFileName, SaveFormat.Docx);
        }

        Console.WriteLine($"Document split into {pageCount} parts successfully.");
    }
}
