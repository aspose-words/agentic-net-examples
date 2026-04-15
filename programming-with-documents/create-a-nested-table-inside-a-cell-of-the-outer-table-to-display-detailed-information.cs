using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create the outer table (3 rows x 4 columns) with placeholder text.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Table");

        // Add the outer table to the document body.
        doc.FirstSection.Body.AppendChild(outerTable);

        // Create the inner table (2 rows x 2 columns) with its own placeholder text.
        Table innerTable = CreateTable(doc, 2, 2, "Inner Table");

        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the resulting document.
        doc.Save("NestedTable.docx");
    }

    // Helper method that builds a table with the specified dimensions and text in each cell.
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

        // Optional metadata for ISO/IEC 29500 compliant documents.
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
