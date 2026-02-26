using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToMarkdown
{
    static void Main()
    {
        // Load the existing Markdown document.
        Document doc = new Document("input.md");

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row belonging to the same document.
            Row newRow = new Row(doc);

            // Determine the number of columns in the existing table.
            int columnCount = table.FirstRow?.Count ?? 0;

            // Add a cell for each column and insert placeholder text.
            for (int i = 0; i < columnCount; i++)
            {
                Cell cell = new Cell(doc);
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Insert sample text into the cell.
                cell.FirstParagraph.AppendChild(new Run(doc, $"New cell {i + 1}"));
                newRow.AppendChild(cell);
            }

            // Append the new row to the end of the table.
            table.AppendChild(newRow);
        }
        else
        {
            Console.WriteLine("No tables found in the document.");
        }

        // Save the modified document back to Markdown format.
        doc.Save("output.md");
    }
}
