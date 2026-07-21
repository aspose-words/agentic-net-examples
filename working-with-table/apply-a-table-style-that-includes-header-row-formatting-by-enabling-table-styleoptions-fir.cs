using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a new table.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Write("Header 1");
        builder.InsertCell();
        builder.Write("Header 2");
        builder.EndRow();

        // Second row – regular data cells.
        builder.InsertCell();
        builder.Write("Data 1");
        builder.InsertCell();
        builder.Write("Data 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Apply a built‑in table style that defines a header row appearance.
        table.StyleIdentifier = StyleIdentifier.MediumShading1Accent1;

        // Enable the FirstRow option so the style is applied to the first (header) row.
        table.StyleOptions = TableStyleOptions.FirstRow;

        // Simple validation that the option was set.
        if ((table.StyleOptions & TableStyleOptions.FirstRow) == 0)
            throw new InvalidOperationException("FirstRow style option was not applied.");

        // Save the document.
        const string outputPath = "TableStyleFirstRow.docx";
        doc.Save(outputPath);
    }
}
