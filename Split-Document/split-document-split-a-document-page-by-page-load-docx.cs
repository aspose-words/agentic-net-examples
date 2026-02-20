using System;
using System.IO;
using Aspose.Words;

class SplitDocumentByPage
{
    static void Main()
    {
        // Load the source DOCX document.
        string sourcePath = @"C:\Docs\Source.docx";
        Document sourceDoc = new Document(sourcePath);

        // Ensure the output directory exists.
        string outputDir = @"C:\Docs\Output";
        Directory.CreateDirectory(outputDir);

        // Loop through each page in the source document.
        for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
        {
            // Extract a single page (pageIndex is zero‑based, count = 1).
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Build a file name for the extracted page (1‑based for readability).
            string outputPath = Path.Combine(outputDir, $"Page_{pageIndex + 1}.docx");

            // Save the extracted page as a separate DOCX file.
            singlePageDoc.Save(outputPath);
        }
    }
}
