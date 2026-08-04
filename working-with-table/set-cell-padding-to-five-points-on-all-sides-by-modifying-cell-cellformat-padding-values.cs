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

        // First row
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Set padding of 5 points on all sides for every cell in the table.
        foreach (Row row in table.Rows)
        {
            foreach (Cell cell in row.Cells)
            {
                // Use SetPaddings to apply left, top, right, bottom padding.
                cell.CellFormat.SetPaddings(5, 5, 5, 5);
            }
        }

        // Save the document to the current directory.
        string outputPath = "CellPadding.docx";
        doc.Save(outputPath);
    }
}
