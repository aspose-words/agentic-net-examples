using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToHtml
{
    static void Main()
    {
        // Load the existing HTML document.
        Document doc = new Document("input.html");

        // Get the first table in the document (adjust index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row belonging to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell (optional, we will add our own cells).
        // newRow.EnsureMinimum();

        // Create cells for the new row.
        // Example: add three cells with sample text.
        for (int i = 1; i <= 3; i++)
        {
            Cell cell = new Cell(doc);
            // Each cell needs a paragraph to hold text.
            Paragraph para = new Paragraph(doc);
            cell.AppendChild(para);
            // Add text to the paragraph.
            Run run = new Run(doc, $"Cell {i} text");
            para.AppendChild(run);
            // Append the cell to the row.
            newRow.AppendChild(cell);
        }

        // Append the new row to the end of the table.
        table.AppendChild(newRow);

        // Save the modified document back to HTML (or another format as needed).
        doc.Save("output.html");
    }
}
