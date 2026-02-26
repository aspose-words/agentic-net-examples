using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableWithHeadersExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // -----------------------------------------------------------------
        // Create the header row. Set HeadingFormat = true so the row repeats
        // at the top of each page when the table spans multiple pages.
        // -----------------------------------------------------------------
        builder.RowFormat.HeadingFormat = true;

        // First header cell.
        builder.InsertCell();
        builder.Write("Product");

        // Second header cell.
        builder.InsertCell();
        builder.Write("Price");

        // End the header row.
        builder.EndRow();

        // -----------------------------------------------------------------
        // Add regular data rows. Disable the heading format for subsequent rows.
        // -----------------------------------------------------------------
        builder.RowFormat.HeadingFormat = false;

        // Example data rows.
        for (int i = 1; i <= 30; i++)
        {
            builder.InsertCell();
            builder.Write($"Item {i}");

            builder.InsertCell();
            builder.Write($"${i * 1.99:F2}");

            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableWithHeaders.docx");
    }
}
