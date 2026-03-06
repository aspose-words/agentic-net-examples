using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Create a new blank document (lifecycle: create)
        Document doc = new Document();

        // Ensure the document has at least one section and a body
        Section section = doc.FirstSection ?? doc.AppendChild(new Section(doc));
        Body body = section.Body ?? section.AppendChild(new Body(doc));

        // Create a table if the document does not already contain one
        Table table = new Table(doc);
        body.AppendChild(table);
        table.EnsureMinimum(); // guarantees at least one row and one cell

        // Create a new row that will be added to the table
        Row newRow = new Row(doc);
        // Ensure the new row has at least one cell before adding content
        newRow.EnsureMinimum();

        // Add a paragraph with some text to the first cell of the new row
        Cell firstCell = newRow.FirstCell;
        Paragraph para = new Paragraph(doc);
        firstCell.FirstParagraph?.RemoveAllChildren(); // clear placeholder paragraph
        firstCell.AppendChild(para);
        para.AppendChild(new Run(doc, "This is a newly added row."));

        // Append the new row to the end of the table
        table.AppendChild(newRow);

        // Save the document (lifecycle: save)
        doc.Save("AddedRow.docx");
    }
}
