using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Load an existing WORDML (or DOCX) document.
        Document doc = new Document("input.docx");

        // Assume the document contains at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Create the first cell for the new row.
        Cell cell1 = new Cell(doc);
        // Add a paragraph with some text to the cell.
        Paragraph para1 = new Paragraph(doc);
        para1.AppendChild(new Run(doc, "New cell 1"));
        cell1.AppendChild(para1);
        // Add the cell to the row.
        newRow.Cells.Add(cell1);

        // Create the second cell for the new row.
        Cell cell2 = new Cell(doc);
        Paragraph para2 = new Paragraph(doc);
        para2.AppendChild(new Run(doc, "New cell 2"));
        cell2.AppendChild(para2);
        newRow.Cells.Add(cell2);

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
