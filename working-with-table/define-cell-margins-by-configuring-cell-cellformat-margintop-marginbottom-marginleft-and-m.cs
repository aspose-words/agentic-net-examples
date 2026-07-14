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

        // Start a table.
        Table table = builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Define cell margins (implemented as padding in Aspose.Words).
        builder.CellFormat.TopPadding = 10;      // Equivalent to MarginTop
        builder.CellFormat.BottomPadding = 10;   // Equivalent to MarginBottom
        builder.CellFormat.LeftPadding = 15;     // Equivalent to MarginLeft
        builder.CellFormat.RightPadding = 15;    // Equivalent to MarginRight

        // Add some text to the cell.
        builder.Write("Cell with custom margins.");

        // Insert a second cell with default margins.
        builder.InsertCell();
        builder.Write("Cell with default margins.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to the current directory.
        const string outputFile = "CellMargins.docx";
        doc.Save(outputFile);
    }
}
