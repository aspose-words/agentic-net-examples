using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Attach a DocumentBuilder to the document – the builder provides a convenient API for inserting content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the Table node that has just been created.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                 // First cell of the first row.
        builder.Write("Cell 1,1");            // Insert text into the cell.

        builder.InsertCell();                 // Second cell of the first row.
        builder.Write("Cell 1,2");

        builder.EndRow();                     // Finish the first row.

        // ---- Second row ----
        builder.InsertCell();                 // First cell of the second row.
        builder.Write("Cell 2,1");

        builder.InsertCell();                 // Second cell of the second row.
        builder.Write("Cell 2,2");

        builder.EndRow();                     // Finish the second row.

        // End the table construction.
        builder.EndTable();

        // Save the document in the legacy DOC format.
        doc.Save("Table.doc");
    }
}
