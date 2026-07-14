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

        // Start a table.
        Table table = builder.StartTable();

        // Build a simple 2x2 table.
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }

        // Finish the table.
        builder.EndTable();

        // Set padding of every cell to 3 points on all sides.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            cell.CellFormat.LeftPadding = 3;
            cell.CellFormat.RightPadding = 3;
            cell.CellFormat.TopPadding = 3;
            cell.CellFormat.BottomPadding = 3;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellPadding.docx");
        doc.Save(outputPath);
    }
}
