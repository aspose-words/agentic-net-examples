using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Build the outer table (3 rows x 4 columns) with placeholder text.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
        // Add the outer table to the document body.
        doc.FirstSection.Body.AppendChild(outerTable);

        // Build the inner (nested) table (2 rows x 2 columns) with its own placeholder text.
        Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to the local file system.
        doc.Save("NestedTable.docx");
    }

    /// <summary>
    /// Creates a table with the specified number of rows and columns.
    /// Each cell contains a single paragraph with the provided text.
    /// </summary>
    private static Table CreateTable(Document doc, int rowCount, int cellCount, string cellText)
    {
        Table table = new Table(doc);

        for (int rowId = 1; rowId <= rowCount; rowId++)
        {
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int cellId = 1; cellId <= cellCount; cellId++)
            {
                Cell cell = new Cell(doc);
                // Ensure the cell has at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add the text to the first paragraph of the cell.
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));

                row.AppendChild(cell);
            }
        }

        // Optional: set title and description for ISO/IEC 29500 compliance.
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
