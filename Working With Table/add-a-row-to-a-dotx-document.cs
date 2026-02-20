using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Get the first table in the document, or create a new one if none exist.
        Table table;
        if (doc.FirstSection.Body.Tables.Count > 0)
        {
            table = doc.FirstSection.Body.Tables[0];
        }
        else
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            table = builder.StartTable();
            builder.InsertCell();
            builder.Write("Header");
            builder.EndRow();
            builder.EndTable();
        }

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // First cell of the new row.
        Cell cell1 = new Cell(doc);
        cell1.AppendChild(new Paragraph(doc));
        cell1.FirstParagraph.AppendChild(new Run(doc, "New Cell 1"));
        newRow.Cells.Add(cell1);

        // Second cell of the new row.
        Cell cell2 = new Cell(doc);
        cell2.AppendChild(new Paragraph(doc));
        cell2.FirstParagraph.AppendChild(new Run(doc, "New Cell 2"));
        newRow.Cells.Add(cell2);

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
