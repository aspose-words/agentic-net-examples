using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Apply vertical text orientation to all cells that will be created.
        builder.CellFormat.Orientation = TextOrientation.Upward;

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Verify that every cell has the vertical orientation applied.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                if (cell.CellFormat.Orientation != TextOrientation.Upward)
                    throw new InvalidOperationException("Cell orientation was not set to vertical.");
            }
        }

        // Save the document to a local folder.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);
        string outputPath = Path.Combine(artifactsDir, "VerticalTable.docx");
        doc.Save(outputPath);
    }
}
