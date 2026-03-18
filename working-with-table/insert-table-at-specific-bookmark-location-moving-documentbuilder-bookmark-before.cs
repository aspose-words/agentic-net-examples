using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableAtBookmark
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that will serve as the insertion point for the table.
        builder.StartBookmark("TableLocation");
        builder.Write("Placeholder text before the table. ");
        builder.EndBookmark("TableLocation");

        // Move the builder's cursor to the bookmark.
        // The cursor will be positioned just after the start of the bookmark,
        // which is the correct place to insert new content.
        bool moved = builder.MoveToBookmark("TableLocation");
        if (!moved)
        {
            throw new InvalidOperationException("Bookmark 'TableLocation' not found.");
        }

        // Build a simple 2x2 table at the bookmark location.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableAtBookmark.docx");
    }
}
