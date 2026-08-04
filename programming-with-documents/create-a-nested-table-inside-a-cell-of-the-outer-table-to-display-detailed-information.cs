using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a blank document.
        Document doc = new Document();

        // Create the outer table (3 rows x 4 columns) with placeholder text.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
        // Add the outer table to the document body.
        doc.FirstSection.Body.AppendChild(outerTable);

        // Create the inner table (2 rows x 2 columns) with placeholder text.
        Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the resulting document.
        doc.Save("NestedTable.docx");
    }

    // Helper method to create a table with the specified dimensions and cell text.
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
                // Ensure each cell contains at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add the provided text to the cell's first paragraph.
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));

                row.AppendChild(cell);
            }
        }

        // Optional metadata for the table.
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
