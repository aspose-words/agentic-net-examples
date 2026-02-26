using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToDocm
{
    static void Main()
    {
        // Load an existing DOCM document.
        Document doc = new Document("InputDocument.docm");

        // Assume the document contains at least one table.
        // Get the first table in the first section.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row belonging to the same document.
        Row newRow = new Row(doc);

        // Ensure the new row has at least one cell.
        newRow.EnsureMinimum();

        // Add content to the first cell of the new row.
        Cell firstCell = newRow.FirstCell;
        Paragraph para = new Paragraph(doc);
        firstCell.FirstParagraph?.Remove(); // Remove the empty paragraph created by EnsureMinimum.
        firstCell.AppendChild(para);
        para.AppendChild(new Run(doc, "New row added programmatically."));

        // Insert the new row after the last existing row of the table.
        table.LastRow.InsertAfter(newRow, table.LastRow);

        // Save the modified document as a DOCM file.
        doc.Save("OutputDocument.docm");
    }
}
