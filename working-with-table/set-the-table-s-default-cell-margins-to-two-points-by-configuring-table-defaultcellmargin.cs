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

        // Get the created table.
        Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

        // Set default cell margins (padding) to 2 points on all sides.
        // Aspose.Words does not expose a Table.DefaultCellMargin property in recent versions,
        // so we apply the padding to each cell individually.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.LeftPadding = 2.0;
                cell.CellFormat.RightPadding = 2.0;
                cell.CellFormat.TopPadding = 2.0;
                cell.CellFormat.BottomPadding = 2.0;
            }
        }

        // Save the document.
        string outputPath = "TableDefaultCellMargin.docx";
        doc.Save(outputPath);

        // Verify the file was saved.
        if (!File.Exists(outputPath))
            throw new Exception("The output document was not created.");

        Console.WriteLine($"Document saved to: {Path.GetFullPath(outputPath)}");
    }
}
