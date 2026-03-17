using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin building a table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // Header row – this row will repeat on every page.
        // -------------------------------------------------
        builder.RowFormat.HeadingFormat = true;                 // Enable repeat‑header flag.
        builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        builder.CellFormat.Width = 100;                         // Width for header cells.

        // First header cell.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.EndRow();

        // Second header cell (same header row).
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // -------------------------------------------------
        // Normal rows – regular data rows.
        // -------------------------------------------------
        builder.RowFormat.HeadingFormat = false;                // Disable repeat‑header for following rows.
        builder.CellFormat.Width = 50;                          // Width for data cells.
        builder.ParagraphFormat.ClearFormatting();              // Reset paragraph formatting.

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

        // Save the document to disk.
        doc.Save("TableWithRepeatingHeader.docx");
    }
}
