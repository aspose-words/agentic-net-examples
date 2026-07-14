using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Build the outer table (3 rows x 4 columns) and add it to the document.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
        doc.FirstSection.Body.AppendChild(outerTable);

        // Build the inner table (2 rows x 2 columns) and insert it into the first cell of the outer table.
        Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to the current directory.
        doc.Save("NestedTable.docx");
    }

    // Helper method that creates a table with the specified dimensions and fills each cell with the given text.
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
                cell.AppendChild(new Paragraph(doc));
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                row.AppendChild(cell);
            }
        }

        // Optional: set title and description for the table (useful for ISO/IEC 29500 compliance).
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
