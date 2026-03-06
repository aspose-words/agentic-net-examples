using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Associate a DocumentBuilder with the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the created Table node.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Cell 1,1");                // Add text to the cell.
        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Cell 1,2");
        builder.EndRow();                         // End the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Cell 2,1");
        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Cell 2,2");
        builder.EndRow();                         // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the document in WORDML (Word 2003 XML) format.
        doc.Save("Table.doc", SaveFormat.WordML);
    }
}
