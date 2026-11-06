using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // Insert three cells with some text.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Cell {i + 1}");
        }

        // End the first row.
        builder.EndRow();

        // Insert a second row with the same number of cells.
        for (int i = 0; i < 3; i++)
        {
            builder.InsertCell();
            builder.Writeln($"Second row, cell {i + 1}");
        }

        // End the second row and the table.
        builder.EndRow();
        builder.EndTable();

        // Retrieve the table we just created.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set vertical alignment of every cell to Bottom.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.VerticalAlignment = CellVerticalAlignment.Bottom;
            }
        }

        // Save the document.
        string outputPath = Path.Combine(Path.GetTempPath(), "VerticalAlignmentBottom.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The document was not saved correctly.");

        // Reload the document and verify one cell's vertical alignment.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChild(NodeType.Table, 0, true);
        Cell firstCell = loadedTable.Rows[0].Cells[0];
        if (firstCell.CellFormat.VerticalAlignment != CellVerticalAlignment.Bottom)
            throw new InvalidOperationException("Cell vertical alignment was not set to Bottom.");

        // Clean up (optional): delete the file after verification.
        // File.Delete(outputPath);
    }
}
