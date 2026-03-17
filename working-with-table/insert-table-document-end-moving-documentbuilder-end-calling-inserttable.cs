using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Move the builder's cursor to the end of the document.
        builder.MoveToDocumentEnd();

        // Start a new table at the current cursor position.
        Table table = builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                     // First cell of the first row.
        builder.Write("Row 1, Cell 1");           // Add text to the cell.
        builder.InsertCell();                     // Second cell of the first row.
        builder.Write("Row 1, Cell 2");
        builder.EndRow();                         // Finish the first row.

        // ---- Second row ----
        builder.InsertCell();                     // First cell of the second row.
        builder.Write("Row 2, Cell 1");
        builder.InsertCell();                     // Second cell of the second row.
        builder.Write("Row 2, Cell 2");
        builder.EndTable();                       // End the table (automatically ends the row).

        // Save the document with the table appended at the end.
        doc.Save("TableAtEnd.docx");
    }
}
