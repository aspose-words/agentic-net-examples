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

        // First cell – set custom margins (implemented as padding).
        builder.InsertCell();
        builder.CellFormat.TopPadding = 10;      // MarginTop
        builder.CellFormat.BottomPadding = 15;   // MarginBottom
        builder.CellFormat.LeftPadding = 20;     // MarginLeft
        builder.CellFormat.RightPadding = 25;    // MarginRight
        builder.Write("Cell with custom margins.");

        // Second cell – keep default margins.
        builder.InsertCell();
        builder.Write("Cell with default margins.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CellMargins.docx");
        doc.Save(outputPath);
    }
}
