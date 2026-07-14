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

        // Start a new table.
        Table table = builder.StartTable();

        // First row – short content.
        builder.RowFormat.HeightRule = HeightRule.Auto; // Ensure auto height.
        builder.InsertCell();
        builder.Write("Short text.");
        builder.InsertCell();
        builder.Write("More short text.");
        builder.EndRow();

        // Second row – longer content to demonstrate auto expansion.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("This is a longer piece of text that should cause the row to expand automatically based on its content.");
        builder.InsertCell();
        builder.Write("Another long text cell that will also cause the row to grow in height automatically.");
        builder.EndRow();

        // Third row – mixed content.
        builder.RowFormat.HeightRule = HeightRule.Auto;
        builder.InsertCell();
        builder.Write("Short.");
        builder.InsertCell();
        builder.Write("A very long text that spans multiple lines and forces the row to increase its height to accommodate all the text without truncation.");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Explicitly set HeightRule to Auto for all rows (in case any were missed).
        foreach (Row row in table.Rows)
        {
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "AutoFitRows.docx");
        doc.Save(outputPath);
    }
}
