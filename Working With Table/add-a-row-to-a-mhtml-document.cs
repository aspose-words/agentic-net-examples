using System;
using Aspose.Words;
using Aspose.Words.Tables;

class AddRowToMhtml
{
    static void Main()
    {
        // Paths to the source MHTML file and the destination file.
        string inputPath = "input.mhtml";
        string outputPath = "output.mhtml";

        // Load the existing MHTML document.
        Document doc = new Document(inputPath);

        // Ensure the document contains at least one table.
        if (doc.FirstSection?.Body?.Tables?.Count > 0)
        {
            // Get the first table in the document.
            Table table = doc.FirstSection.Body.Tables[0];

            // Create a new row that belongs to the same document.
            Row newRow = new Row(doc);

            // Determine how many cells (columns) the new row should have.
            int columnCount = table.FirstRow?.Cells?.Count ?? 1;

            // Populate the new row with cells.
            for (int i = 0; i < columnCount; i++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Each cell must contain at least one paragraph.
                Paragraph para = new Paragraph(doc);
                cell.AppendChild(para);

                // Add some sample text to the paragraph.
                Run run = new Run(doc, $"New cell {i + 1}");
                para.AppendChild(run);

                // Append the cell to the new row.
                newRow.AppendChild(cell);
            }

            // Append the completed row to the end of the table.
            table.AppendChild(newRow);
        }
        else
        {
            Console.WriteLine("No tables found in the document.");
        }

        // Save the modified document back to MHTML format.
        doc.Save(outputPath, SaveFormat.Mhtml);
    }
}
