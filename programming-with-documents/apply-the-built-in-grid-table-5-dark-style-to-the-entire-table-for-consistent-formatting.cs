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

        // Build a simple 2x3 table.
        Table table = builder.StartTable();

        // First row (header).
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        // Third row.
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply the built‑in "Grid Table 5 Dark" style to the entire table.
        table.StyleIdentifier = StyleIdentifier.GridTable5Dark;
        // Optional: enable additional style options if desired.
        table.StyleOptions = TableStyleOptions.FirstColumn | TableStyleOptions.LastColumn | TableStyleOptions.RowBands;

        // Save the document to disk.
        const string outputPath = "StyledTable.docx";
        doc.Save(outputPath);
    }
}
