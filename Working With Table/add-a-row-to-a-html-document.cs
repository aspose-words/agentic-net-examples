using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
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

            // Create a new row associated with the document.
            Row newRow = new Row(doc);

            // For each existing column, add a new cell to the row.
            // Here we assume the table has at least one row to determine column count.
            int columnCount = table.FirstRow?.Cells?.Count ?? 1;
            for (int i = 0; i < columnCount; i++)
            {
                // Create a new cell.
                Cell cell = new Cell(doc);

                // Add a paragraph with some text to the cell.
                Paragraph para = new Paragraph(doc);
                Run run = new Run(doc, $"New cell {i + 1}");
                para.AppendChild(run);
                cell.AppendChild(para);

                // Add the cell to the new row.
                newRow.Cells.Add(cell);
            }

            // Append the new row to the table.
            table.Rows.Add(newRow);
        }

        // Save the modified document back to HTML.
        doc.Save("output.html", SaveFormat.Html);
    }
}
