using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct the outer table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table outerTable = builder.StartTable();

        // First row of the outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 1");
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.EndRow();

        // Second row of the outer table.
        builder.InsertCell();
        builder.Write("Outer Cell 3");
        builder.InsertCell();
        builder.Write("Outer Cell 4");
        builder.EndTable(); // Finish the outer table.

        // Create an inner table (2 rows x 2 columns) with sample text.
        Table innerTable = CreateTable(doc, 2, 2, "Inner");

        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to the local file system.
        doc.Save("NestedTable.docx");
    }

    // Helper method that builds a table with the specified dimensions and text.
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
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add the provided text to the paragraph.
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                row.AppendChild(cell);
            }
        }

        // Optional: set title/description for completeness.
        table.Title = "Nested Table";
        table.Description = "A table inserted inside a cell of another table.";

        return table;
    }
}
