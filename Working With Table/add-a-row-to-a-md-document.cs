using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for easier content insertion.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and keep a reference to it.
        Table table = builder.StartTable();

        // Build the first row using the builder (two cells with sample text).
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // ----- Add a new row using the Row class -----
        // Create a new Row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell.
        newRow.EnsureMinimum();

        // Populate the first cell of the new row.
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New cell 1"));

        // Create a second cell, add a paragraph and a run with text.
        Cell secondCell = new Cell(doc);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "New cell 2"));

        // Append the second cell to the new row.
        newRow.AppendChild(secondCell);

        // Append the completed row to the existing table.
        table.AppendChild(newRow);
        // ---------------------------------------------

        // Finish the table.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("AddedRow.docx");
    }
}
