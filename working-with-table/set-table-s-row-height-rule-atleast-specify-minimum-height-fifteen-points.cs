using System;
using Aspose.Words;
using Aspose.Words.Tables;

class SetRowHeightExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // Add a single cell with some text.
        builder.InsertCell();
        builder.Write("Sample cell");
        builder.EndRow(); // Finish the first (and only) row.

        // Access the RowFormat of the created row.
        RowFormat rowFormat = table.FirstRow.RowFormat;

        // Set the height rule to AtLeast and specify a minimum height of 15 points.
        rowFormat.HeightRule = HeightRule.AtLeast;
        rowFormat.Height = 15.0;

        // Save the document to disk.
        doc.Save("RowHeightAtLeast.docx");
    }
}
