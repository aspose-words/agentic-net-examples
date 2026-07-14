using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ----- Header row (repeated on each page) -----
        builder.RowFormat.HeadingFormat = true;               // Mark this row as a heading.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Width = 100;                       // Width for header cells.

        // First header cell.
        builder.InsertCell();
        builder.Write("Header Column 1");
        // Second header cell.
        builder.InsertCell();
        builder.Write("Header Column 2");
        builder.EndRow();                                     // End header row.

        // ----- Data rows (no longer heading) -----
        builder.CellFormat.Width = 50;                        // Width for data cells.
        builder.ParagraphFormat.ClearFormatting();           // Reset paragraph formatting.
        builder.RowFormat.HeadingFormat = false;             // Ensure following rows are not headings.

        // Add enough rows to make the table span multiple pages.
        for (int i = 0; i < 50; i++)
        {
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 1");
            builder.InsertCell();
            builder.Write($"Row {i + 1}, Column 2");
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        doc.Save("TableWithRepeatingHeader.docx");
    }
}
