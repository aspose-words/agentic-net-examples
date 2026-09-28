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

        // Begin a table.
        builder.StartTable();

        // First cell.
        builder.InsertCell();
        Cell cell = builder.CurrentParagraph.ParentNode as Cell;
        // Set padding to 5 points on all sides.
        cell.CellFormat.LeftPadding = 5;
        cell.CellFormat.RightPadding = 5;
        cell.CellFormat.TopPadding = 5;
        cell.CellFormat.BottomPadding = 5;
        builder.Writeln("Cell 1");

        // Second cell.
        builder.InsertCell();
        cell = builder.CurrentParagraph.ParentNode as Cell;
        // Set padding to 5 points on all sides.
        cell.CellFormat.LeftPadding = 5;
        cell.CellFormat.RightPadding = 5;
        cell.CellFormat.TopPadding = 5;
        cell.CellFormat.BottomPadding = 5;
        builder.Writeln("Cell 2");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        const string outputPath = "TableWithPadding.docx";
        doc.Save(outputPath);

        // Simple validation that the file was created.
        if (!System.IO.File.Exists(outputPath))
            throw new Exception("Failed to create the output document.");
    }
}
