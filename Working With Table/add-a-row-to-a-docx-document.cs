using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("Input.docx");

        // Get the first table in the document (adjust index as needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row for the document.
        Row newRow = new Row(doc);

        // Add cells to the new row. Here we add two cells with sample text.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "New cell 1"));
        newRow.Cells.Add(cell1);

        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "New cell 2"));
        newRow.Cells.Add(cell2);

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
