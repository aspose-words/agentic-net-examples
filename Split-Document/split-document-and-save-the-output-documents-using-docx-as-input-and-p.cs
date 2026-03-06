using System;
using Aspose.Words;

class SplitDocxToPdf
{
    static void Main()
    {
        // Path to the source DOCX file.
        string inputPath = @"C:\Docs\SourceDocument.docx";

        // Load the DOCX document.
        Document sourceDoc = new Document(inputPath);

        // Iterate through each page of the source document.
        for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
        {
            // Extract a single page as a new Document.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Build the output PDF file name (e.g., SourceDocument_Page1.pdf).
            string outputPath = $@"C:\Docs\SourceDocument_Page{pageIndex + 1}.pdf";

            // Save the extracted page as PDF. The format is inferred from the .pdf extension.
            pageDoc.Save(outputPath);
        }
    }
}
