using System;
using Aspose.Words;
using Aspose.Words.Tables;

class TableWrapExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Set the table's text wrapping to "Around", which corresponds to a square wrap.
        table.TextWrapping = TextWrapping.Around;

        // Optional: define the distance between the table and surrounding text.
        table.AbsoluteHorizontalDistance = 10; // points
        table.AbsoluteVerticalDistance = 10;   // points

        // Save the document.
        doc.Save("TableWrapSquare.docx");
    }
}
