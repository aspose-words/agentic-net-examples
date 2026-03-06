using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToDocm
{
    static void Main()
    {
        // Load the existing DOCM document.
        Document doc = new Document("input.docm");

        // Get the first table in the document (adjust index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Ensure the row has at least one cell.
        newRow.EnsureMinimum();

        // Add some text to the first cell of the new row.
        Cell firstCell = newRow.FirstCell;
        firstCell.FirstParagraph.AppendChild(new Run(doc, "Text in the new row"));

        // Save the modified document as a DOCM file.
        doc.Save("output.docm");
    }
}
