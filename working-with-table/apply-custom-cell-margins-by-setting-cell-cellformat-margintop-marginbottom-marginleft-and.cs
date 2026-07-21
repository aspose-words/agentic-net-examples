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

        // Insert the first cell and apply custom margins (implemented via padding properties).
        builder.InsertCell();
        builder.CellFormat.TopPadding = 10;    // Equivalent to MarginTop
        builder.CellFormat.BottomPadding = 12; // Equivalent to MarginBottom
        builder.CellFormat.LeftPadding = 8;   // Equivalent to MarginLeft
        builder.CellFormat.RightPadding = 8;  // Equivalent to MarginRight
        builder.Write("Cell 1 with custom margins.");

        // Insert a second cell with default margins.
        builder.InsertCell();
        builder.Write("Cell 2 with default margins.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomCellMargins.docx");
        doc.Save(outputPath);
    }
}
