using System;
using Aspose.Words;
using Aspose.Words.Tables;

class CellPaddingExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // Insert the first cell.
        builder.InsertCell();

        // Set padding of 3 points on all sides for the current cell.
        // Using the SetPaddings method (left, top, right, bottom).
        builder.CellFormat.SetPaddings(3, 3, 3, 3);

        // Add some text to the cell.
        builder.Write("Cell with 3‑point padding.");

        // End the row and the table.
        builder.EndRow();
        builder.EndTable();

        // Save the document.
        doc.Save("CellPadding.docx");
    }
}
