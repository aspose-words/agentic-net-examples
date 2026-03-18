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

        // Begin a new table.
        Table table = builder.StartTable();

        // ---------- First Row ----------
        // Insert a cell and add some text.
        builder.InsertCell();
        builder.Write("First row, first cell.");

        // End the first row.
        builder.EndRow();

        // Set the first row's height rule to Auto.
        // No explicit Height is assigned, so the row will grow automatically to fit its contents.
        Row firstRow = table.FirstRow;
        firstRow.RowFormat.HeightRule = HeightRule.Auto;

        // ---------- Second Row ----------
        // Insert a cell for the second row.
        builder.InsertCell();

        // For demonstration, set a fixed height on this row.
        builder.RowFormat.Height = 50;               // Height in points.
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        builder.Write("Second row, first cell.");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Save the document to a file.
        doc.Save("RowHeightAuto.docx");
    }
}
