using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Fields;

class BookmarkTableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a table with 5 rows. Each row will have its own bookmark.
        builder.StartTable();

        for (int i = 1; i <= 5; i++)
        {
            // Insert a single cell for the row.
            builder.InsertCell();

            // Create a bookmark that encloses the cell's text.
            string bookmarkName = $"Row{i}";
            builder.StartBookmark(bookmarkName);
            builder.Write($"This is the content of row {i}.");
            builder.EndBookmark(bookmarkName);

            // End the current row.
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Add a blank paragraph to separate the table from the list of links.
        builder.Writeln();

        // Insert hyperlinks that navigate to each row's bookmark.
        for (int i = 1; i <= 5; i++)
        {
            // Format the hyperlink text.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;

            // Insert a hyperlink field that points to the bookmark.
            string displayText = $"Go to Row {i}";
            string bookmarkTarget = $"Row{i}";
            FieldHyperlink link = (FieldHyperlink)builder.InsertHyperlink(displayText, bookmarkTarget, true);
            link.ScreenTip = $"Jump to row {i}";

            // Move to the next line after each hyperlink.
            builder.Writeln();
        }

        // Save the document to disk.
        doc.Save("BookmarksInTable.docx");
    }
}
