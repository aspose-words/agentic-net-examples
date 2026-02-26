using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToHtml
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row that belongs to the same document.
            Row newRow = new Row(doc);

            // Append the new row to the end of the table.
            table.AppendChild(newRow);

            // Ensure the row has at least one cell.
            newRow.EnsureMinimum();

            // Add a second cell to the row (optional, for a 2‑column table).
            Cell secondCell = new Cell(doc);
            secondCell.AppendChild(new Paragraph(doc));
            newRow.AppendChild(secondCell);

            // Populate the first cell with text.
            newRow.FirstCell.FirstParagraph.AppendChild(new Run(doc, "New cell 1"));

            // Populate the second cell with text.
            secondCell.FirstParagraph.AppendChild(new Run(doc, "New cell 2"));
        }

        // Save the modified document back to HTML.
        doc.Save("output.html");
    }
}
