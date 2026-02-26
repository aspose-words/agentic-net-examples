using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToDotx
{
    static void Main()
    {
        // Load the existing DOTX template.
        Document doc = new Document("Template.dotx");

        // Find the first table in the document (adjust as needed for your scenario).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell.
        newRow.EnsureMinimum();

        // Add content to the first cell of the new row.
        Cell firstCell = newRow.FirstCell;
        Paragraph para = new Paragraph(doc);
        firstCell.AppendChild(para);
        Run run = new Run(doc, "New row added programmatically.");
        para.AppendChild(run);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Save the modified document back as a DOTX (or any other format you need).
        doc.Save("ModifiedTemplate.dotx");
    }
}
