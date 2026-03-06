using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

class DocumentProcessing
{
    static void Main()
    {
        // Paths to the source and auxiliary documents.
        string sourcePath = @"input.docx";
        string insertPath = @"insert.docx";
        string appendPath = @"append.docx";

        // Output paths.
        string clonedPath = @"cloned.docx";
        string insertedPath = @"inserted.docx";
        string appendedPath = @"appended.docx";
        string splitPart1Path = @"split_part1.docx";
        string splitPart2Path = @"split_part2.docx";

        // Load the original document.
        Document original = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 1. Clone the original document (deep copy).
        // -----------------------------------------------------------------
        Document cloned = original.Clone();
        cloned.Save(clonedPath); // Save the cloned document.

        // -----------------------------------------------------------------
        // 2. Insert an additional document into the cloned document.
        // -----------------------------------------------------------------
        Document docToInsert = new Document(insertPath);
        DocumentBuilder insertBuilder = new DocumentBuilder(cloned);
        // Move the cursor to the beginning of the document.
        insertBuilder.MoveToDocumentStart();
        // Insert the whole document at the cursor position.
        insertBuilder.InsertDocument(docToInsert, ImportFormatMode.KeepSourceFormatting);
        cloned.Save(insertedPath); // Save the document after insertion.

        // -----------------------------------------------------------------
        // 3. Append another document to the end of the original document.
        // -----------------------------------------------------------------
        Document docToAppend = new Document(appendPath);
        original.AppendDocument(docToAppend, ImportFormatMode.KeepSourceFormatting);
        original.Save(appendedPath); // Save the document after appending.

        // -----------------------------------------------------------------
        // 4. Split the original (now appended) document into two parts.
        // -----------------------------------------------------------------
        // Ensure the document has at least two pages.
        if (original.PageCount < 2)
            throw new InvalidOperationException("Document does not contain enough pages to split.");

        // Determine a split point (e.g., halfway).
        int splitPageIndex = original.PageCount / 2; // Zero‑based index of the first page of the second part.

        // Extract pages for the first part (pages 0 .. splitPageIndex‑1).
        Document part1 = original.ExtractPages(0, splitPageIndex - 1);
        part1.Save(splitPart1Path);

        // Extract pages for the second part (pages splitPageIndex .. last).
        Document part2 = original.ExtractPages(splitPageIndex, original.PageCount - 1);
        part2.Save(splitPart2Path);
    }
}
