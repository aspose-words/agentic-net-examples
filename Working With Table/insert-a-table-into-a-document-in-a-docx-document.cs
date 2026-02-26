using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table. The builder returns the created Table object.
        Table table = builder.StartTable();

        // ---- First row ----
        // Insert first cell and add text.
        builder.InsertCell();
        builder.Write("Cell 1, Row 1");

        // Insert second cell in the same row and add text.
        builder.InsertCell();
        builder.Write("Cell 2, Row 1");

        // End the first row.
        builder.EndRow();

        // ---- Second row ----
        builder.InsertCell();
        builder.Write("Cell 1, Row 2");

        builder.InsertCell();
        builder.Write("Cell 2, Row 2");

        // End the second row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as a DOCX file.
        doc.Save("TableInserted.docx");
    }
}
