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

        // Create the outer table (e.g., 3 rows x 4 columns) and add it to the document.
        Table outerTable = CreateTable(doc, 3, 4, "Outer Table");
        doc.FirstSection.Body.AppendChild(outerTable);

        // Create the inner table (e.g., 2 rows x 2 columns) and insert it into the first cell of the outer table.
        Table innerTable = CreateTable(doc, 2, 2, "Inner Table");
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Define the output path and ensure the directory exists.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedTable.docx");
        doc.Save(outputPath);
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
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                // Add the specified text to the first paragraph of the cell.
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));

                row.AppendChild(cell);
            }
        }

        // Optional: set title and description for better accessibility.
        table.Title = "Aspose table title";
        table.Description = "Aspose table description";

        return table;
    }
}
