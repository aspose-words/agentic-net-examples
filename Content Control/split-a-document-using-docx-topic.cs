using System;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentSplitter
{
    static void Main()
    {
        // Load the source DOCX document.
        // The constructor automatically detects the file format.
        Document sourceDoc = new Document("InputDocument.docx");

        // Determine the total number of pages in the document.
        int totalPages = sourceDoc.PageCount;

        // Loop through each page and extract it as a separate document.
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // ExtractPages returns a new Document containing the specified page range.
            Document pageDoc = sourceDoc.ExtractPages(pageIndex, pageIndex);

            // Save each extracted page as an individual DOCX file.
            // The file name includes the page number for clarity.
            string outputFileName = $"Output_Page_{pageIndex}.docx";
            pageDoc.Save(outputFileName, SaveFormat.Docx);
        }

        Console.WriteLine("Document split into individual pages successfully.");
    }
}
