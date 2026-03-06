using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Path to the source DOCX document.
        string sourcePath = @"C:\Docs\SourceDocument.docx";

        // Load the source document (DOCX format is detected automatically).
        Document sourceDoc = new Document(sourcePath);

        // Determine the number of pages in the document.
        // GetPageInfo is zero‑based; the last page index is PageCount‑1.
        int pageCount = sourceDoc.PageCount;

        // Split the document: extract each page into a separate document and save it.
        for (int page = 1; page <= pageCount; page++)
        {
            // Extract a single page (pages are 1‑based for ExtractPages).
            Document pageDoc = sourceDoc.ExtractPages(page, page);

            // Save each page as a separate DOCX file.
            string outPath = $@"C:\Docs\Page_{page}.docx";
            pageDoc.Save(outPath, SaveFormat.Docx);
        }

        // -----------------------------------------------------------------
        // List all supported input (load) formats.
        Console.WriteLine("Supported input formats (LoadFormat):");
        foreach (LoadFormat loadFmt in Enum.GetValues(typeof(LoadFormat)))
        {
            // Convert the LoadFormat to its typical file extension.
            string ext = FileFormatUtil.LoadFormatToExtension(loadFmt);
            Console.WriteLine($"- {loadFmt} ({ext})");
        }

        // -----------------------------------------------------------------
        // List all supported output (save) formats.
        Console.WriteLine("\nSupported output formats (SaveFormat):");
        foreach (SaveFormat saveFmt in Enum.GetValues(typeof(SaveFormat)))
        {
            // Not all SaveFormat values have a corresponding extension; handle exceptions.
            try
            {
                string ext = FileFormatUtil.SaveFormatToExtension(saveFmt);
                Console.WriteLine($"- {saveFmt} ({ext})");
            }
            catch (ArgumentException)
            {
                // Skip formats that cannot be mapped to an extension.
                Console.WriteLine($"- {saveFmt} (no extension mapping)");
            }
        }
    }
}
