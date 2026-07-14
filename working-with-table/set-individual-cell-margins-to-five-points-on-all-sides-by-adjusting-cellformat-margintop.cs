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

        // Start a table and populate it with 2 rows and 2 columns.
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

        // Finish the table.
        builder.EndTable();

        // Set the padding (margins) of each cell to 5 points on all sides.
        foreach (Row r in table.Rows)
        {
            foreach (Cell c in r.Cells)
            {
                c.CellFormat.TopPadding = 5;
                c.CellFormat.BottomPadding = 5;
                c.CellFormat.LeftPadding = 5;
                c.CellFormat.RightPadding = 5;
            }
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellMargins.docx");
        doc.Save(outputPath);
    }
}
