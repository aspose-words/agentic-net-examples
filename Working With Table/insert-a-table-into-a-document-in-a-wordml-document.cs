using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which provides a convenient way to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the Table node that was created.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Row 1, Cell 1");           // Insert text into the cell.
        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Row 1, Cell 2");
        builder.EndRow();                         // Finish the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Row 2, Cell 2");
        builder.EndRow();                         // Finish the second row.

        // End the table construction.
        builder.EndTable();

        // Optional: set table metadata.
        table.Title = "Sample Table";
        table.Description = "Demonstrates inserting a table into a WORDML document.";

        // Save the document in WORDML format (XML representation of a Word document).
        doc.Save("TableInsert.doc", SaveFormat.WordML);
    }
}
