using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ParagraphManipulation
{
    static void Main()
    {
        // Load the original document from disk.
        Document original = new Document("Original.docx");

        // Create a deep copy of the original document.
        Document cloned = original.Clone();

        // Create an empty document that will hold inserted and appended content.
        Document container = new Document();

        // Use DocumentBuilder to position the cursor at the end of the container.
        DocumentBuilder builder = new DocumentBuilder(container);
        builder.MoveToDocumentEnd();

        // Insert a page break before inserting the cloned document (optional formatting).
        builder.InsertBreak(BreakType.PageBreak);

        // Insert the cloned document at the current cursor position.
        builder.InsertDocument(cloned, ImportFormatMode.KeepSourceFormatting);

        // Load a second document that will be appended to the container.
        Document second = new Document("Second.docx");

        // Append the second document to the end of the container.
        container.AppendDocument(second, ImportFormatMode.KeepSourceFormatting);

        // Save the combined document.
        container.Save("Combined.docx");

        // Split the combined document into two parts:
        // Part 1 – pages 1 to 2.
        Document part1 = container.ExtractPages(1, 2);

        // Part 2 – remaining pages (from page 3 to the end).
        int remainingPages = container.PageCount - 2;
        Document part2 = container.ExtractPages(3, remainingPages);

        // Save the split documents.
        part1.Save("Combined_Part1.docx");
        part2.Save("Combined_Part2.docx");
    }
}
