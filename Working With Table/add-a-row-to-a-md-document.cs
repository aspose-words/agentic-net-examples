using System;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

class AddRowToMarkdownTable
{
    static void Main()
    {
        // Load an existing Markdown document that contains at least one table.
        Document doc = new Document("input.md", new LoadOptions { LoadFormat = LoadFormat.Markdown });

        // Get the first table in the document.
        Table table = doc.FirstSection.Body.Tables[0];

        // Create a new row for the document.
        Row newRow = new Row(doc);

        // Determine how many cells the new row should have (match existing rows).
        int cellCount = table.FirstRow.Cells.Count;

        // Populate the new row with cells and sample text.
        for (int i = 0; i < cellCount; i++)
        {
            // Create a new cell.
            Cell cell = new Cell(doc);

            // Create a paragraph and a run with the desired text.
            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc, $"New cell {i + 1}");

            // Assemble the paragraph hierarchy.
            para.AppendChild(run);
            cell.AppendChild(para);

            // Add the cell to the new row.
            newRow.Cells.Add(cell);
        }

        // Append the new row to the end of the table.
        table.Rows.Add(newRow);

        // Save the modified document back to Markdown format.
        doc.Save("output.md", SaveFormat.Markdown);
    }
}
