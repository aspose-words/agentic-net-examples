using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class SetTableLeftIndent
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Initialize a DocumentBuilder for the document.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table and insert the first cell.
        Table table = builder.StartTable();
        builder.InsertCell();

        // Write some sample text into the cell.
        builder.Write("Table with left indent of 2 cm");

        // End the current row and the table.
        builder.EndRow();
        builder.EndTable();

        // Convert 2 centimeters to points (1 cm = 72 / 2.54 points).
        double twoCentimetersInPoints = 2 * 72.0 / 2.54; // ≈ 56.6929 points

        // Apply the left indent to the table.
        table.LeftIndent = twoCentimetersInPoints;

        // Save the document to a file.
        doc.Save("TableWithLeftIndent.docx");
    }
}
