using System;
using Aspose.Words;
using Aspose.Words.Tables;

class GenerateTableWithUniqueBookmarks
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Define the number of rows you want in the table.
        int rowCount = 5;

        for (int i = 0; i < rowCount; i++)
        {
            // Generate a unique bookmark name for the current row.
            string bookmarkName = $"RowBookmark_{i + 1}";

            // Insert a bookmark start before the row content.
            builder.StartBookmark(bookmarkName);

            // Insert cells for the row. Adjust the number of cells as needed.
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 1");

            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 2");

            // End the bookmark after the row content.
            builder.EndBookmark(bookmarkName);

            // End the current row.
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableWithUniqueBookmarks.docx");
    }
}
