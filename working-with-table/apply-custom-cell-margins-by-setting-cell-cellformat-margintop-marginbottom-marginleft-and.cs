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

        // ---------- First Row ----------
        // First cell.
        builder.InsertCell();
        Cell cell1 = table.FirstRow.FirstCell;
        // Set custom cell paddings (margins) in points.
        cell1.CellFormat.TopPadding = 10;
        cell1.CellFormat.BottomPadding = 10;
        cell1.CellFormat.LeftPadding = 5;
        cell1.CellFormat.RightPadding = 5;
        builder.Write("Cell 1");

        // Second cell.
        builder.InsertCell();
        Cell cell2 = table.FirstRow.LastCell;
        cell2.CellFormat.TopPadding = 15;
        cell2.CellFormat.BottomPadding = 15;
        cell2.CellFormat.LeftPadding = 8;
        cell2.CellFormat.RightPadding = 8;
        builder.Write("Cell 2");

        builder.EndRow();

        // ---------- Second Row ----------
        // First cell.
        builder.InsertCell();
        Cell cell3 = table.LastRow.FirstCell;
        cell3.CellFormat.TopPadding = 12;
        cell3.CellFormat.BottomPadding = 12;
        cell3.CellFormat.LeftPadding = 6;
        cell3.CellFormat.RightPadding = 6;
        builder.Write("Cell 3");

        // Second cell.
        builder.InsertCell();
        Cell cell4 = table.LastRow.LastCell;
        cell4.CellFormat.TopPadding = 20;
        cell4.CellFormat.BottomPadding = 20;
        cell4.CellFormat.LeftPadding = 10;
        cell4.CellFormat.RightPadding = 10;
        builder.Write("Cell 4");

        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomCellMargins.docx");
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new Exception("The document was not saved correctly.");
    }
}
