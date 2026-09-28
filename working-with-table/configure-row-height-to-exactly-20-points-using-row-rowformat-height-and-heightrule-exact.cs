using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple table with two rows.
        builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Writeln("First row, first cell");
        builder.EndRow();

        // Second row – this row will have its height set to exactly 20 points.
        builder.InsertCell();
        builder.Writeln("Second row, first cell");
        builder.EndRow();

        // End the table.
        builder.EndTable();

        // Access the second row (index 1) and set its height.
        Row secondRow = doc.FirstSection.Body.Tables[0].Rows[1];
        secondRow.RowFormat.Height = 20; // Height in points.
        secondRow.RowFormat.HeightRule = HeightRule.Exactly; // Exact height rule.

        // Save the document.
        string outputPath = "RowHeightExact.docx";
        doc.Save(outputPath);

        // Verify that the file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("Failed to create the output document.");
    }
}
