using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToPdf
{
    static void Main()
    {
        // Create a new blank Word document.
        Document doc = new Document();

        // Initialize a DocumentBuilder to construct the document content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table and add an initial row with one cell.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Original Row, Cell 1");
        builder.EndRow();

        // Retrieve the table that was just created.
        Table table = (Table)builder.CurrentParagraph.ParentNode;

        // Create a new row that will be added to the table.
        Row newRow = new Row(doc);

        // Create a cell for the new row and add a paragraph with text.
        Cell newCell = new Cell(doc);
        newCell.AppendChild(new Paragraph(doc));
        newCell.FirstParagraph.AppendChild(new Run(doc, "New Row, Cell 1"));
        newRow.AppendChild(newCell);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // End the table construction.
        builder.EndTable();

        // Save the document as a PDF file.
        doc.Save("AddedRow.pdf", SaveFormat.Pdf);
    }
}
