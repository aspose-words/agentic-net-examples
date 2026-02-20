using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class ApplyTableStyleOptionsToMarkdown
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // Insert the first row (header).
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Insert a data row.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style.
        table.StyleIdentifier = StyleIdentifier.LightGrid;

        // Apply specific style options using the TableStyleOptions flags.
        // Here we enable formatting for the first row and row banding.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Optional: let the table auto‑fit its contents.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a Markdown file.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("TableWithStyle.md", saveOptions);
    }
}
