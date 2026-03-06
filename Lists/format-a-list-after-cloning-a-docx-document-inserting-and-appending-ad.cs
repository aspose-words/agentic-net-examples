using System;
using Aspose.Words;
using Aspose.Words.Lists;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the original DOCX document.
        Document srcDoc = new Document("Original.docx");

        // Create a deep clone of the source document.
        Document clonedDoc = srcDoc.Clone(true) as Document;

        // Insert additional content into the cloned document.
        DocumentBuilder builder = new DocumentBuilder(clonedDoc);
        builder.MoveToDocumentEnd();
        builder.InsertBreak(BreakType.PageBreak);
        builder.Writeln("Additional content inserted into the cloned document.");

        // Prepare import options to preserve list numbering from the cloned part.
        ImportFormatOptions importOptions = new ImportFormatOptions
        {
            KeepSourceNumbering = true
        };

        // Append the cloned document to the original document.
        srcDoc.AppendDocument(clonedDoc, ImportFormatMode.KeepSourceFormatting, importOptions);

        // Update list labels so numbering is correct after the append.
        srcDoc.UpdateListLabels();

        // Save the combined document.
        srcDoc.Save("Combined.docx");

        // Split the combined document: extract the first two pages.
        Document firstPart = srcDoc.ExtractPages(1, 2);
        firstPart.Save("Combined_Part1.docx");

        // Extract the remaining pages (from page 3 to the end).
        Document secondPart = srcDoc.ExtractPages(3, srcDoc.PageCount);
        secondPart.Save("Combined_Part2.docx");
    }
}
