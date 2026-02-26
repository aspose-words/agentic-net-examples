using System;
using Aspose.Words;
using Aspose.Words.Saving;

class VbaMacroDemo
{
    static void Main()
    {
        // Path to the original document that may contain VBA macros.
        string sourcePath = @"C:\Docs\SourceWithMacros.docx";

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Check if the document has macros.
        bool hasMacros = sourceDoc.HasMacros;
        Console.WriteLine($"Source document has macros: {hasMacros}");

        // Clone the loaded document (deep copy).
        Document clonedDoc = (Document)sourceDoc.Clone(true);

        // -----------------------------------------------------------------
        // Insert another document into the cloned document at the current cursor.
        // -----------------------------------------------------------------
        string insertPath = @"C:\Docs\InsertDoc.docx";
        Document insertDoc = new Document(insertPath);

        // Use DocumentBuilder to position the cursor and insert the document.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.MoveToDocumentEnd(); // Move cursor to the end of the cloned document.
        builder.InsertDocument(insertDoc, ImportFormatMode.KeepSourceFormatting);

        // -----------------------------------------------------------------
        // Append a third document to the end of the cloned document.
        // -----------------------------------------------------------------
        string appendPath = @"C:\Docs\AppendDoc.docx";
        Document appendDoc = new Document(appendPath);
        clonedDoc.AppendDocument(appendDoc, ImportFormatMode.UseDestinationStyles);

        // -----------------------------------------------------------------
        // Save the combined document.
        // -----------------------------------------------------------------
        string combinedPath = @"C:\Docs\CombinedResult.docx";
        clonedDoc.Save(combinedPath);

        // -----------------------------------------------------------------
        // Split the combined document into individual pages.
        // Each page is saved as a separate DOCX file.
        // -----------------------------------------------------------------
        int totalPages = clonedDoc.PageCount;
        for (int pageIndex = 1; pageIndex <= totalPages; pageIndex++)
        {
            // ExtractPages uses 1‑based page numbers.
            Document pageDoc = clonedDoc.ExtractPages(pageIndex, pageIndex);
            string pagePath = $@"C:\Docs\Page_{pageIndex}.docx";
            pageDoc.Save(pagePath);
        }

        // -----------------------------------------------------------------
        // Optional: Remove macros from the combined document and save.
        // -----------------------------------------------------------------
        if (clonedDoc.HasMacros)
        {
            clonedDoc.RemoveMacros();
            string noMacroPath = @"C:\Docs\Combined_NoMacros.docx";
            clonedDoc.Save(noMacroPath);
        }

        Console.WriteLine("Processing completed.");
    }
}
