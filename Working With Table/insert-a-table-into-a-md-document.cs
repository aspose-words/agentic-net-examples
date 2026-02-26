using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document for easy content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder's cursor is now inside the table.
        Table table = builder.StartTable();

        // ---------- First row (header) ----------
        builder.InsertCell();               // First cell of the row.
        builder.Write("Header 1");          // Insert text into the cell.
        builder.InsertCell();               // Second cell of the row.
        builder.Write("Header 2");
        builder.EndRow();                   // Complete the first row.

        // ---------- Second row ----------
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();                   // Complete the second row.

        // Finish the table.
        builder.EndTable();

        // Configure Markdown save options (optional: set alignment, etc.).
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            // Example: align all table contents to the left when exporting.
            // TableContentAlignment = TableContentAlignment.Left
        };

        // Save the document as a Markdown file. The table will be exported in Markdown format.
        doc.Save("TableDocument.md", saveOptions);
    }
}
