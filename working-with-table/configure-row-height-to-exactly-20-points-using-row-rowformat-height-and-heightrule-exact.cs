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

        // Set the height of the next row to exactly 20 points.
        builder.RowFormat.Height = 20;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Second row – will have the exact height defined above.
        builder.InsertCell();
        builder.Write("Second row (height = 20 points, Exact).");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document to the current directory.
        doc.Save("RowHeightExact.docx");
    }
}
