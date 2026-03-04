using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class IterateDocumentPages
{
    static void Main()
    {
        // Load the source DOCX document.
        // The Document constructor is the approved way to create/load a document.
        Document sourceDoc = new Document("InputDocument.docx");

        // Ensure the layout is up‑to‑date so that page information is accurate.
        sourceDoc.UpdatePageLayout();

        // Iterate through each page using the page count.
        // For each page we extract a single‑page document using ExtractPages
        // and then save it as an individual DOCX file.
        for (int pageIndex = 0; pageIndex < sourceDoc.PageCount; pageIndex++)
        {
            // Extract a document that contains only the current page.
            // ExtractPages uses zero‑based page index and a count of 1.
            Document singlePageDoc = sourceDoc.ExtractPages(pageIndex, 1);

            // Build a file name for the extracted page.
            string outFileName = $"Page_{pageIndex + 1}.docx";

            // Save the single‑page document.
            // The Save(string) overload automatically determines the format from the extension.
            singlePageDoc.Save(outFileName);
        }

        Console.WriteLine("Pages have been extracted and saved as separate DOCX files.");
    }
}
