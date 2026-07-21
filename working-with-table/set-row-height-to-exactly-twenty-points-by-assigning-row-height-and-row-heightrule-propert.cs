using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class RowHeightExample
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

        // Second row – set exact height of 20 points.
        builder.RowFormat.Height = 20.0;               // Height in points.
        builder.RowFormat.HeightRule = HeightRule.Exactly; // Exact height rule.
        builder.InsertCell();
        builder.Write("Second row, cell 1.");
        builder.InsertCell();
        builder.Write("Second row, cell 2.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Define output file path.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "RowHeightExample.docx");

        // Save the document.
        doc.Save(outputPath);

        // Reload the document to verify the row height.
        Document loadedDoc = new Document(outputPath);
        Table loadedTable = loadedDoc.FirstSection.Body.Tables[0];
        Row secondRow = loadedTable.Rows[1];

        // Validate that the height and rule are as expected.
        if (Math.Abs(secondRow.RowFormat.Height - 20.0) > 0.01 ||
            secondRow.RowFormat.HeightRule != HeightRule.Exactly)
        {
            throw new Exception("Row height was not set to exactly 20 points.");
        }

        // Indicate successful completion.
        Console.WriteLine("Document created and row height verified successfully.");
    }
}
