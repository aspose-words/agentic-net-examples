using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the outer table (e.g., 3 rows x 4 columns) and add it to the document.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Cell");
        doc.FirstSection.Body.AppendChild(outerTable);

        // Build the inner (nested) table (e.g., 2 rows x 2 columns).
        Table innerTable = CreateTable(doc, 2, 2, "Inner Cell");

        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to the current directory.
        string outputPath = Path.Combine(Environment.CurrentDirectory, "NestedTable.docx");
        doc.Save(outputPath);
    }

    // Helper method that creates a table with the specified dimensions and fills each cell with the given text.
    private static Table CreateTable(Document doc, int rowCount, int columnCount, string cellText)
    {
        Table table = new Table(doc);

        for (int rowIdx = 1; rowIdx <= rowCount; rowIdx++)
        {
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int colIdx = 1; colIdx <= columnCount; colIdx++)
            {
                Cell cell = new Cell(doc);
                // Ensure the cell contains at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add the specified text to the cell.
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                row.AppendChild(cell);
            }
        }

        // Optional: set title/description for better accessibility (ignored in older formats).
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
