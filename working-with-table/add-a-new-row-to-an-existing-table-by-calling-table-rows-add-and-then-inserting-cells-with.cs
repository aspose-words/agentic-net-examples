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

        // Create a table and add it to the document body.
        Table table = new Table(doc);
        doc.FirstSection.Body.AppendChild(table);

        // Build an initial row with two cells.
        Row firstRow = new Row(doc);
        table.Rows.Add(firstRow);

        // First cell of the initial row.
        Cell firstCell = new Cell(doc);
        firstCell.AppendChild(new Paragraph(doc));
        firstCell.FirstParagraph.AppendChild(new Run(doc, "First row, cell 1"));
        firstRow.Cells.Add(firstCell);

        // Second cell of the initial row.
        Cell secondCell = new Cell(doc);
        secondCell.AppendChild(new Paragraph(doc));
        secondCell.FirstParagraph.AppendChild(new Run(doc, "First row, cell 2"));
        firstRow.Cells.Add(secondCell);

        // Add a new row to the existing table.
        Row newRow = new Row(doc);
        table.Rows.Add(newRow);

        // Insert cells into the new row.
        Cell newCell1 = new Cell(doc);
        newCell1.AppendChild(new Paragraph(doc));
        newCell1.FirstParagraph.AppendChild(new Run(doc, "New row, cell 1"));
        newRow.Cells.Add(newCell1);

        Cell newCell2 = new Cell(doc);
        newCell2.AppendChild(new Paragraph(doc));
        newCell2.FirstParagraph.AppendChild(new Run(doc, "New row, cell 2"));
        newRow.Cells.Add(newCell2);

        // Optional validation: ensure the table now has two rows.
        if (table.Rows.Count != 2)
            throw new InvalidOperationException("The table should contain exactly two rows.");

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Result.docx");
        doc.Save(outputPath);
    }
}
