using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Load the DOTX template.
        Document doc = new Document("Template.dotx");

        // Assume the document contains at least one table.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row belonging to the same document.
        Row newRow = new Row(doc);

        // Ensure the row has at least one cell (required before adding content).
        newRow.EnsureMinimum();

        // Determine how many cells the existing rows have.
        int columnCount = table.FirstRow.Cells.Count;

        // Remove the automatically created cell so we can add the correct number of cells.
        newRow.RemoveAllChildren();

        // Add cells to the new row.
        for (int i = 0; i < columnCount; i++)
        {
            Cell cell = new Cell(doc);
            // Each cell must contain at least one paragraph.
            cell.AppendChild(new Paragraph(doc));
            // Add some sample text to the cell.
            cell.FirstParagraph.AppendChild(new Run(doc, $"New cell {i + 1}"));
            newRow.AppendChild(cell);
        }

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document.
        doc.Save("Result.docx");
    }
}
