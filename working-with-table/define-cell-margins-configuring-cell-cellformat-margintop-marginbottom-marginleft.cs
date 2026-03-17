using System;
using Aspose.Words;
using Aspose.Words.Tables;

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Begin a table.
        builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Define cell margins by setting the padding properties.
        // These paddings act as the space between the cell borders and its contents.
        builder.CellFormat.LeftPadding = 12;   // left margin (points)
        builder.CellFormat.TopPadding = 8;     // top margin (points)
        builder.CellFormat.RightPadding = 12;  // right margin (points)
        builder.CellFormat.BottomPadding = 8;  // bottom margin (points)

        // Add some text to the cell.
        builder.Write("Cell with custom margins.");

        // Insert a second cell with default padding.
        builder.InsertCell();
        builder.Write("Second cell.");

        // Finish the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document to disk.
        doc.Save("CellMargins.docx");
    }
}
