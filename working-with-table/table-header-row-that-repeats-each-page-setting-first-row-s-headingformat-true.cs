using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ----- Header row (repeated on each page) -----
        // Set the HeadingFormat flag so this row repeats on every page.
        builder.RowFormat.HeadingFormat = true;

        // Optional: center the text in the header cells.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        // Define a reasonable width for the header cells.
        builder.CellFormat.Width = 100;

        // First header cell.
        builder.InsertCell();
        builder.Write("Header 1");

        // Second header cell.
        builder.InsertCell();
        builder.Write("Header 2");

        // End the header row.
        builder.EndRow();

        // ----- Data rows (regular rows) -----
        // Turn off the heading flag for subsequent rows.
        builder.RowFormat.HeadingFormat = false;

        // Reset any paragraph formatting that was applied to the header.
        builder.ParagraphFormat.ClearFormatting();

        // Set a narrower width for data cells.
        builder.CellFormat.Width = 50;

        // Add enough rows to force the table to span multiple pages.
        for (int i = 0; i < 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 1");

            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 2");

            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableWithRepeatingHeader.docx");
    }
}
