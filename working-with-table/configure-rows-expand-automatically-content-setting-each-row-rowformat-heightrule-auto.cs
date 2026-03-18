using System;
using Aspose.Words;
using Aspose.Words.Tables;

class ConfigureRowHeightAuto
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Use DocumentBuilder to construct a simple table.
        DocumentBuilder builder = new DocumentBuilder(doc);
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("This is a long piece of text that should cause the row to expand automatically.");
        builder.InsertCell();
        builder.Write("Short");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Another long text that will test the auto‑height behavior of the row.");
        builder.InsertCell();
        builder.Write("More text");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Ensure every row in the table grows automatically to fit its content.
        foreach (Row row in table.Rows)
        {
            // Set the HeightRule to Auto for each row.
            row.RowFormat.HeightRule = HeightRule.Auto;
        }

        // Save the document to a file.
        string outputPath = "ConfiguredRows.docx";
        doc.Save(outputPath);
        Console.WriteLine($"Document saved to {outputPath}");
    }
}
