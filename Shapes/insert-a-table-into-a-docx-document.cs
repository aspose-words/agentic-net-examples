using System;
using Aspose.Words;
using Aspose.Words.Tables;

class InsertTableExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder which will be used to insert content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Write("Header 1");
        // First row, second cell.
        builder.InsertCell();
        builder.Write("Header 2");
        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        // Second row, second cell.
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        // End the second row.
        builder.EndRow();

        // Third row, first cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 1");
        // Third row, second cell.
        builder.InsertCell();
        builder.Write("Row 2, Cell 2");
        // End the third row.
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to a DOCX file.
        doc.Save("InsertedTable.docx");
    }
}
