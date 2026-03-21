using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocxPageSplitter
{
    static void Main()
    {
        // Use folders relative to the executable location
        string baseDir = AppContext.BaseDirectory;
        string sourceFolder = Path.Combine(baseDir, "Input");
        string outputFolder = Path.Combine(baseDir, "Output");

        // Ensure both directories exist
        Directory.CreateDirectory(sourceFolder);
        Directory.CreateDirectory(outputFolder);

        // If there are no .docx files, create a simple sample document
        string[] docxFiles = Directory.GetFiles(sourceFolder, "*.docx");
        if (docxFiles.Length == 0)
        {
            string samplePath = Path.Combine(sourceFolder, "Sample.docx");
            var sampleDoc = new Document();
            var builder = new DocumentBuilder(sampleDoc);
            builder.Writeln("This is page 1.");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("This is page 2.");
            sampleDoc.Save(samplePath);
            docxFiles = new[] { samplePath };
        }

        // Process each .docx file in the source folder
        foreach (string docxPath in docxFiles)
        {
            // Load the document (load rule)
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

        Console.WriteLine("Processing complete. PDFs are located in:");
        Console.WriteLine(outputFolder);
    }
}
