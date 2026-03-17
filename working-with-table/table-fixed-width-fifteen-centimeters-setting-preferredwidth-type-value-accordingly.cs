using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetTableFixedWidth
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert a single cell with some sample text.
        builder.InsertCell();
        builder.Write("Sample cell");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Convert 15 centimeters to points (1 inch = 2.54 cm, 1 point = 1/72 inch).
        double points = 15.0 * 72.0 / 2.54; // ≈ 425.19685 points

        // Set the table's preferred width to the calculated points value.
        table.PreferredWidth = PreferredWidth.FromPoints(points);

        // Save the document to disk.
        doc.Save("TableWithFixedWidth.docx");
    }
}
