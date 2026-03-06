using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document
        Document doc = new Document();

        // Create a new table and add it to the document body
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Ensure the table has at least one row and one cell
        table.EnsureMinimum();

        // Add content to the initial cell so the table is not empty
        table.FirstRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "Initial cell"));

        // Create a new row that will be added to the table
        Row newRow = new Row(doc);

        // Append the new row to the table (after the existing rows)
        table.AppendChild(newRow);

        // Ensure the new row has at least one cell
        newRow.EnsureMinimum();

        // Add text to the first cell of the new row
        newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New row, first cell"));

        // Optionally add a second cell to the new row
        Cell secondCell = new Cell(doc);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "New row, second cell"));
        newRow.AppendChild(secondCell);

        // Save the document to a DOCX file
        doc.Save("AddedRow.docx");
    }
}
