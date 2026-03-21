using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowExample
{
    static void Main()
    {
        // Create a new document and add a simple table with one row.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start the table.
        builder.StartTable();

        // Add two cells to the first row.
        builder.InsertCell();
        builder.Write("Existing cell 1");
        builder.InsertCell();
        builder.Write("Existing cell 2");

        // End the first row.
        builder.EndRow();

        // Retrieve the table we just created.
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
