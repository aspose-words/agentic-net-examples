using System;
using Aspose.Words;
using Aspose.Words.Tables;

class RenderTableWithHeaders
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // ---------- Header row ----------
        // Insert first cell of the header row.
        builder.InsertCell();
        // Apply header formatting (bold) and set the cell text.
        builder.Font.Bold = true;
        builder.Writeln("Product");
        // Insert second cell of the header row.
        builder.InsertCell();
        builder.Writeln("Quantity");
        // End the header row.
        builder.EndRow();

        // Mark the first row as a repeating header.
        // In Aspose.Words the HeadingFormat flag is sufficient – the row will repeat on each page.
        table.FirstRow.RowFormat.HeadingFormat = true;
        // Removed the non‑existent RepeatHeader property which caused the compile error.

        // Reset font style for data rows.
        builder.Font.Bold = false;

        // ---------- Data rows ----------
        // Add a few sample data rows.
        for (int i = 1; i <= 10; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Item {i}");
            builder.InsertCell();
            builder.Writeln((i * 5).ToString());
            builder.EndRow();
        }

        // End the table.
        builder.EndTable();

        // Save the document to disk.
        doc.Save("TableWithHeaders.docx");
    }
}
