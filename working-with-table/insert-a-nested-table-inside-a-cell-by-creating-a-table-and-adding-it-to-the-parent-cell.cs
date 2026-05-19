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
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build the outer table (2 rows x 2 columns) using the DocumentBuilder.
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
        builder.EndTable();

        // Create the inner table (2 rows x 2 columns) using a helper method.
        Table innerTable = CreateTable(doc, 2, 2, "Inner");

        // Insert the inner table into the first cell of the outer table.
        outerTable.FirstRow.FirstCell.AppendChild(innerTable);

        // Save the document to the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "NestedTable.docx");
        doc.Save(outputPath);
    }

    // Helper method that creates a table with the specified dimensions and text.
    private static Table CreateTable(Document doc, int rowCount, int cellCount, string cellText)
    {
        Table table = new Table(doc);

        for (int r = 0; r < rowCount; r++)
        {
            Row row = new Row(doc);
            table.AppendChild(row);

            for (int c = 0; c < cellCount; c++)
            {
                Cell cell = new Cell(doc);
                // Each cell must contain at least one paragraph.
                cell.AppendChild(new Paragraph(doc));
                cell.FirstParagraph.AppendChild(new Run(doc, cellText));
                row.AppendChild(cell);
            }
        }

        return table;
    }
}
