using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        builder.StartTable();

        builder.InsertCell();
        builder.Writeln("Cell 1");
        builder.InsertCell();
        builder.Writeln("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Writeln("Cell 3");
        builder.InsertCell();
        builder.Writeln("Cell 4");
        builder.EndRow();

        builder.EndTable();

        // Retrieve the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);
        if (table == null)
            throw new InvalidOperationException("Table was not created.");

        // Apply background shading to every cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.Shading.BackgroundPatternColor = Color.LightGray;
            }
        }

        // Save the document.
        string outputPath = "TableShading.docx";
        doc.Save(outputPath);

        // Verify the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not saved.", outputPath);
    }
}
