using System;
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
        Table table = builder.StartTable();

        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // Set padding of 3 points on all sides for every cell in the table.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            cell.CellFormat.LeftPadding = 3;
            cell.CellFormat.RightPadding = 3;
            cell.CellFormat.TopPadding = 3;
            cell.CellFormat.BottomPadding = 3;
        }

        // Save the document.
        doc.Save("CellPadding.docx");
    }
}
