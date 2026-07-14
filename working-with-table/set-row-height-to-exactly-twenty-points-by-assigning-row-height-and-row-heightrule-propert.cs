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

        // First row (default height).
        builder.InsertCell();
        builder.Write("First row, cell 1.");
        builder.InsertCell();
        builder.Write("First row, cell 2.");
        builder.EndRow();

        // Configure the next row to have an exact height of 20 points.
        builder.RowFormat.Height = 20;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second row with the specified height.
        builder.InsertCell();
        builder.Write("Second row, cell 1.");
        builder.InsertCell();
        builder.Write("Second row, cell 2.");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Save the document.
        doc.Save("RowHeightExample.docx");
    }
}
