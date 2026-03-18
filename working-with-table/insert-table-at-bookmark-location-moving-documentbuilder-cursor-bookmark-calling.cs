using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document and associate a DocumentBuilder with it.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a bookmark that will mark the insertion point for the table.
        builder.StartBookmark("InsertTableHere");
        builder.Write("Placeholder text before the table. ");
        builder.EndBookmark("InsertTableHere");

        // Move the builder's cursor to the bookmark so that subsequent inserts occur there.
        builder.MoveToBookmark("InsertTableHere");

        // Start a table at the current cursor position.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();               // First cell of the first row.
        builder.Write("Row 1, Cell 1");

        builder.InsertCell();               // Second cell of the first row.
        builder.Write("Row 1, Cell 2");
        builder.EndRow();                   // End the first row.

        // ---- Second row ----
        builder.InsertCell();               // First cell of the second row.
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();               // Second cell of the second row.
        builder.Write("Row 2, Cell 2");
        builder.EndRow();                   // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the resulting document.
        doc.Save("TableAtBookmark.docx");
    }
}
