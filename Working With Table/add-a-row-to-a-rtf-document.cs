using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to start building content.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table with a single row containing two cells.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow(); // First row completed.

        // ----- Add a new row programmatically -----
        // Create a new Row instance attached to the same document.
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell cell1 = new Cell(doc);
        // Add a paragraph and a run with text to the cell.
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "New Cell 1"));
        cell1.AppendChild(para1);
        // Add the cell to the row.
        newRow.Cells.Add(cell1);

        // Second cell of the new row.
        Cell cell2 = new Cell(doc);
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "New Cell 2"));
        cell2.AppendChild(para2);
        newRow.Cells.Add(cell2);

        // Insert the new row at the end of the table.
        table.Rows.Add(newRow);
        // -------------------------------------------

        // Finish the table.
        builder.EndTable();

        // Save the document as RTF using RtfSaveOptions.
        RtfSaveOptions rtfOptions = new RtfSaveOptions();
        doc.Save("AddedRow.rtf", rtfOptions);
    }
}
