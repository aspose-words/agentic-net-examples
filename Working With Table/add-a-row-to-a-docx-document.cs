using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Create a table and add it to the document body.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Ensure the table has at least one row and one cell.
        table.EnsureMinimum();

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the new row has at least one cell (required for a valid row).
        newRow.EnsureMinimum();

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Add a paragraph with some text to the first cell of the new row.
        Cell firstCell = newRow.FirstCell;
        Paragraph paragraph = new Paragraph(doc);
        firstCell.AppendChild(paragraph);
        Run run = new Run(doc, "This is a newly added row.");
        paragraph.AppendChild(run);

        // Save the document to a DOCX file.
        doc.Save("AddedRow.docx");
    }
}
