using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxPageSplitter
{
    // Entry point
    static void Main()
    {
        // Folder containing the source DOCX files
        string sourceFolder = @"C:\Docs\Input";
        // Folder where the per‑page PDFs will be written
        string outputFolder = @"C:\Docs\Output";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Process each .docx file in the source folder
        foreach (string docxPath in Directory.GetFiles(sourceFolder, "*.docx"))
        {
            // Load the document from file (load rule)
            Document sourceDoc = new Document(docxPath);

            // Determine the number of pages in the document
            int pageCount = sourceDoc.PageCount;

            // Base name without extension for naming output files
            string baseName = Path.GetFileNameWithoutExtension(docxPath);

            // Extract each page and save it as an individual PDF
            for (int pageIndex = 0; pageIndex < pageCount; pageIndex++)
            {
                // Extract a single page (pageIndex is zero‑based)
                Document pageDoc = sourceDoc.ExtractPages(pageIndex, 1);

                // Build the output PDF file name
                string pdfPath = Path.Combine(
                    outputFolder,
                    $"{baseName}_Page_{pageIndex + 1}.pdf");

                // Save the extracted page as PDF (save rule)
                pageDoc.Save(pdfPath, SaveFormat.Pdf);
            }
        }
    }
}
