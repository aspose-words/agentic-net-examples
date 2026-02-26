using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new Word document (will be saved as PDF later)
        Document doc = new Document();

        // Use DocumentBuilder to start a table
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.StartTable();

        // First row with two cells
        builder.InsertCell();
        builder.Write("Row 1, Cell 1");
        builder.InsertCell();
        builder.Write("Row 1, Cell 2");
        builder.EndRow();

        // Add a new row manually using the Row class
        // Get reference to the table that was just created
        Table table = (Table)builder.CurrentParagraph.ParentNode;

        // Create a new Row instance belonging to the same document
        Row newRow = new Row(doc);

        // Append the new row to the table
        table.AppendChild(newRow);

        // Ensure the row has at least one cell before adding content
        newRow.EnsureMinimum();

        // Fill cells of the new row
        Cell firstCell = newRow.FirstCell;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 1"));

        // Add a second cell to the new row
        Cell secondCell = new Cell(doc);
        newRow.AppendChild(secondCell);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "Row 2, Cell 2"));

        // End the table
        builder.EndTable();

        // Save the document as PDF
        doc.Save("AddedRow.pdf");
    }
}
