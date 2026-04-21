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

        // Start a table and fill it with sample data (2 rows x 2 columns).
        Table table = builder.StartTable();
        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 2; col++)
            {
                builder.InsertCell();
                builder.Write($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Set padding of every cell to 3 points on all sides.
        foreach (Cell cell in table.GetChildNodes(NodeType.Cell, true))
        {
            cell.CellFormat.TopPadding = 3;
            cell.CellFormat.BottomPadding = 3;
            cell.CellFormat.LeftPadding = 3;
            cell.CellFormat.RightPadding = 3;
        }

        // Save the document to disk.
        string outputPath = "CellPadding.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("Document was not saved successfully.");
    }
}
