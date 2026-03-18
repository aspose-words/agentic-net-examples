using System;
using Aspose.Words;

class SplitDocument
{
    static void Main()
    {
        // Load the original document (replace with your actual file path).
        Document source = new Document("input.docx");

        // Total number of pages in the source document.
        int totalPages = source.PageCount;

        // Iterate through each page, extract it, and save as an individual DOCX file.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // Extract a single page (page numbers are 1‑based).
            Document pageDoc = source.ExtractPages(pageIndex, pageIndex);

            // Construct the output file name for this page.
            string outputPath = $"output_page_{pageIndex}.docx";

            // Save the extracted page. The Save method automatically preserves
            // all formatting, styles, and page orientation of the original page.
            pageDoc.Save(outputPath);
        }
    }
}
