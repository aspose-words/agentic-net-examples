using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The method returns the Table node that was created.
        Table table = builder.StartTable();

        // ---- First row ----
        // Insert first cell and write text into it.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");

        // Insert second cell of the first row.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");

        // End the first row.
        builder.EndRow();

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");

        builder.InsertCell();
        builder.Write("Row 2, Cell 2");

        // End the table.
        builder.EndTable();

        // Save the document to a DOCX file.
        doc.Save("TableExample.docx");
    }
}
