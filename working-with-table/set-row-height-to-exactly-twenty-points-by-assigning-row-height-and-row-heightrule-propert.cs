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

        // First row – uses default height settings.
        builder.InsertCell();
        builder.Write("First row (default height).");
        builder.EndRow();

        // Configure the next row to have an exact height of 20 points.
        builder.RowFormat.Height = 20;               // Height in points.
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second row – will be exactly 20 points tall.
        builder.InsertCell();
        builder.Write("Second row (height = 20 points).");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        doc.Save("RowHeightExample.docx");
    }
}
