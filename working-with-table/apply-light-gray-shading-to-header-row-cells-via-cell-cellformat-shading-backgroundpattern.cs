using System;
using System.Drawing;
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

        // Start a new table.
        builder.StartTable();

        // ----- Header row -----
        // First header cell.
        builder.InsertCell();
        builder.Writeln("Header 1");
        // Second header cell.
        builder.InsertCell();
        builder.Writeln("Header 2");
        // End the header row.
        builder.EndRow();

        // ----- Data rows -----
        // First data row.
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 1, Cell 2");
        builder.EndRow();

        // Second data row.
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 1");
        builder.InsertCell();
        builder.Writeln("Row 2, Cell 2");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply light gray shading to each cell in the header row.
        Row headerRow = table.FirstRow;
        foreach (Cell cell in headerRow.Cells)
        {
            cell.CellFormat.Shading.Texture = TextureIndex.TextureSolid;
            cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
        }

        // Save the document.
        string outputPath = "TableWithHeaderShading.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
