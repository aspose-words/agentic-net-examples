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

        // Build an initial 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("R1C1");
        builder.InsertCell();
        builder.Write("R1C2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("R2C1");
        builder.InsertCell();
        builder.Write("R2C2");
        builder.EndRow();

        // Finish the table construction.
        builder.EndTable();

        // Insert a new column at index 1 (between the existing columns).
        // Since Table.InsertColumn does not exist, we add a new cell to each row at the desired position.
        int insertIndex = 1; // zero‑based column index
        foreach (Row row in table.Rows)
        {
            // Create a new empty cell with a paragraph (required for text insertion later).
            Cell newCell = new Cell(doc);
            newCell.AppendChild(new Paragraph(doc));

            // Insert the cell at the specified column index.
            if (insertIndex < row.Cells.Count)
                row.InsertBefore(newCell, row.Cells[insertIndex]);
            else
                row.AppendChild(newCell);
        }

        // Populate the newly inserted column with sample text.
        int rowNumber = 1;
        foreach (Row row in table.Rows)
        {
            Cell newCell = row.Cells[insertIndex];
            newCell.FirstParagraph.AppendChild(new Run(doc, $"NewCol{rowNumber}"));
            rowNumber++;
        }

        // Save the document.
        string outputPath = "InsertedColumn.docx";
        doc.Save(outputPath);

        // Simple validation to ensure the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved correctly.");
    }
}
