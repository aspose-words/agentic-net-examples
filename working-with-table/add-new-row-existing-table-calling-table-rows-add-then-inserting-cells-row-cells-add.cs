using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Load an existing Word document that already contains a table.
        Document doc = new Document("Input.docx");

        // Get the first table in the document (adjust the index if needed).
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row that belongs to the same document.
        Row newRow = new Row(doc);

        // Define the number of cells you want to add to the new row.
        int cellCount = 3; // Example: three columns.

        for (int i = 0; i < cellCount; i++)
        {
            // Create a new cell for the row.
            Cell cell = new Cell(doc);

            // Ensure the cell has at least one paragraph.
            cell.EnsureMinimum();

            // Add some text to the cell's first paragraph.
            cell.FirstParagraph.AppendChild(new Run(doc, $"New cell {i + 1}"));

            // Add the cell to the row's Cells collection.
            newRow.Cells.Add(cell);
        }

        // Append the newly created row to the table's Rows collection.
        table.Rows.Add(newRow);

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
