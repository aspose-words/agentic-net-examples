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

        // Set the default cell margins (padding) to 2 points on all sides for every cell.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                cell.CellFormat.TopPadding = 2;
                cell.CellFormat.BottomPadding = 2;
                cell.CellFormat.LeftPadding = 2;
                cell.CellFormat.RightPadding = 2;
            }
        }

        // Save the document.
        string outputPath = "TableDefaultCellMargin.docx";
        doc.Save(outputPath);

        // Verify that the file was saved.
        if (!File.Exists(outputPath))
            throw new FileNotFoundException("The output document was not created.", outputPath);
    }
}
