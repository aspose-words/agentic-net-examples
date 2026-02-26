using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the table.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // ---- First row ----
        builder.InsertCell();                 // First cell of the first row.
        builder.Write("Cell 1, Row 1");       // Add text to the cell.
        builder.InsertCell();                 // Second cell of the first row.
        builder.Write("Cell 2, Row 1");
        builder.EndRow();                     // End the first row.

        // ---- Second row ----
        builder.InsertCell();                 // First cell of the second row.
        builder.Write("Cell 1, Row 2");
        builder.InsertCell();                 // Second cell of the second row.
        builder.Write("Cell 2, Row 2");
        builder.EndRow();                     // End the second row.

        // Finish the table.
        builder.EndTable();

        // Save the document in RTF format.
        doc.Save("TableInRtf.rtf");
    }
}
