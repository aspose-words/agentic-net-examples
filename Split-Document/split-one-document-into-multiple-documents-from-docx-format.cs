using System;
using System.IO;
using Aspose.Words;

class DocumentSplitter
{
    static void Main()
    {
        // Path to the source DOCX file.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the source document using the Document constructor (lifecycle rule).
        Document sourceDoc = new Document(sourcePath);

        // Determine how many pages the document has.
        int totalPages = sourceDoc.PageCount;

        // Create an output directory for the split documents.
        string outputDir = @"C:\Docs\SplitPages";
        if (!Directory.Exists(outputDir))
            Directory.CreateDirectory(outputDir);

        // Loop through each page, extract it as a separate Document, and save it.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // ExtractPages uses 1‑based page numbers.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Build the output file name.
            string outFile = Path.Combine(outputDir, $"Page_{pageIndex}.docx");

            // Save the extracted page document (lifecycle rule).
            pageDoc.Save(outFile);
        }

        // Optional: split by sections instead of pages.
        // Uncomment the following block to split by sections.
        /*
        string sectionOutputDir = @"C:\Docs\SplitSections";
        if (!Directory.Exists(sectionOutputDir))
            Directory.CreateDirectory(sectionOutputDir);

        for (int secIndex = 0; secIndex < sourceDoc.Sections.Count; secIndex++)
        {
            // Clone the whole document and then remove all sections except the current one.
            Document sectionDoc = (Document)sourceDoc.Clone(true);
            // Remove sections before the target.
            for (int i = secIndex - 1; i >= 0; i--)
                sectionDoc.Sections[i].Remove();
            // Remove sections after the target.
            for (int i = sectionDoc.Sections.Count - 1; i > 0; i--)
                sectionDoc.Sections[i].Remove();

            string outFile = Path.Combine(sectionOutputDir, $"Section_{secIndex + 1}.docx");
            sectionDoc.Save(outFile);
        }
        */
    }
}
