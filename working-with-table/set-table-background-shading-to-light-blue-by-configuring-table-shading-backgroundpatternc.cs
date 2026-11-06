using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a table.
        builder.StartTable();

        // First row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 1");

        // End the first row.
        builder.EndRow();

        // Second row, first cell.
        builder.InsertCell();
        builder.Writeln("Cell 2");

        // End the second row.
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChildNodes(NodeType.Table, true)[0];

        // Apply light blue background shading to every cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightBlue;
                cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            }
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Validate that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not saved.");

        // Re-open the document to verify shading.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = (Table)loadedDoc.GetChildNodes(NodeType.Table, true)[0];

        // Verify that the first cell has the expected shading color.
        Cell firstCell = loadedTable.Rows[0].Cells[0];
        Color bgColor = firstCell.CellFormat.Shading.BackgroundPatternColor;

        if (bgColor.ToArgb() != Color.LightBlue.ToArgb())
            throw new Exception("Table shading color does not match the expected light blue.");

        // Program completed successfully.
        Console.WriteLine("Table created with light blue background shading and saved to " + outputPath);
    }
}
