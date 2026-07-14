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

        // Start a table and add the first row.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("First row, first cell.");
        builder.EndRow();

        // Configure the next row to have an exact height of 20 points.
        builder.RowFormat.Height = 20.0;
        builder.RowFormat.HeightRule = HeightRule.Exactly;

        // Add the second row which will inherit the height settings.
        builder.InsertCell();
        builder.Write("Second row, first cell.");
        builder.EndTable();

        // Optional validation to ensure the height was applied.
        if (table.Rows.Count >= 2)
        {
            Row secondRow = table.Rows[1];
            if (secondRow.RowFormat.Height != 20.0 || secondRow.RowFormat.HeightRule != HeightRule.Exactly)
                throw new InvalidOperationException("Row height was not set correctly.");
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the document.
        string outputPath = Path.Combine(outputDir, "RowHeightExact.docx");
        doc.Save(outputPath);
    }
}
