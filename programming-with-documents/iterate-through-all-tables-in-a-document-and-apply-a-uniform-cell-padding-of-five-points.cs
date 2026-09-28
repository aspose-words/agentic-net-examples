using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a sample table with 2 rows and 3 columns.
        Table table = builder.StartTable();
        for (int col = 0; col < 3; col++)
        {
            builder.InsertCell();
            builder.Writeln($"Header {col + 1}");
        }
        builder.EndRow();

        for (int row = 0; row < 2; row++)
        {
            for (int col = 0; col < 3; col++)
            {
                builder.InsertCell();
                builder.Writeln($"R{row + 1}C{col + 1}");
            }
            builder.EndRow();
        }
        builder.EndTable();

        // Iterate through all tables and set uniform cell padding of five points.
        foreach (Table tbl in doc.GetChildNodes(NodeType.Table, true).OfType<Table>())
        {
            foreach (Row r in tbl.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    // Set uniform padding (points) for the cell.
                    c.CellFormat.LeftPadding = 5;
                    c.CellFormat.RightPadding = 5;
                    c.CellFormat.TopPadding = 5;
                    c.CellFormat.BottomPadding = 5;
                }
            }
        }

        // Save the document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
