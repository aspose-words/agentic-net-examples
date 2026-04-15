using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Settings;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Optimize the document for Microsoft Word 2016.
        doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2016);

        // Build a complex table (outer table with a nested inner table).
        DocumentBuilder builder = new DocumentBuilder(doc);

        // ---------- Outer table ----------
        builder.StartTable();

        // First row, two cells.
        builder.InsertCell();
        builder.Write("Outer Cell 1");
        builder.InsertCell();
        builder.Write("Outer Cell 2");
        builder.EndRow();

        // Second row, first cell will contain a nested table.
        builder.InsertCell();
        // Placeholder text for the cell that will hold the nested table.
        builder.Write("Outer Cell 3 (contains nested table)");
        builder.InsertCell();
        builder.Write("Outer Cell 4");
        builder.EndRow();

        // Finish the outer table.
        builder.EndTable();

        // Retrieve the outer table we just created.
        Table outerTable = (Table)doc.GetChild(NodeType.Table, 0, true);

        // ---------- Nested table ----------
        // Create a new table that will be placed inside the first cell of the outer table.
        Table nestedTable = new Table(doc);

        // Build the nested table (2 rows x 2 columns).
        for (int rowIdx = 0; rowIdx < 2; rowIdx++)
        {
            Row row = new Row(doc);
            nestedTable.AppendChild(row);

            for (int colIdx = 0; colIdx < 2; colIdx++)
            {
                Cell cell = new Cell(doc);
                cell.AppendChild(new Paragraph(doc));
                cell.FirstParagraph.AppendChild(new Run(doc, $"Nested {rowIdx + 1},{colIdx + 1}"));
                row.AppendChild(cell);
            }
        }

        // Optionally set a simple border for the nested table.
        nestedTable.SetBorder(BorderType.Left, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        nestedTable.SetBorder(BorderType.Right, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        nestedTable.SetBorder(BorderType.Top, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);
        nestedTable.SetBorder(BorderType.Bottom, LineStyle.Single, 1.0, System.Drawing.Color.Black, true);

        // Insert the nested table into the first cell of the outer table.
        Cell targetCell = outerTable.FirstRow.FirstCell;
        // Remove the placeholder paragraph that contains the earlier text.
        targetCell.RemoveAllChildren();
        targetCell.AppendChild(nestedTable);

        // Save the document to the local file system.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComplexTable.docx");
        doc.Save(outputPath);
    }
}
