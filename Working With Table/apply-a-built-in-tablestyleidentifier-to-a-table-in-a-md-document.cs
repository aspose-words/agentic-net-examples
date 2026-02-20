using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class ApplyTableStyleToMarkdown
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a simple 2x2 table.
        Table table = builder.StartTable();

        // First row – header cells.
        builder.InsertCell();
        builder.Writeln("Header 1");
        builder.InsertCell();
        builder.Writeln("Header 2");
        builder.EndRow();

        // Second row – data cells.
        builder.InsertCell();
        builder.Writeln("Data 1");
        builder.InsertCell();
        builder.Writeln("Data 2");
        builder.EndRow();

        // End the table construction.
        builder.EndTable();

        // Apply a built‑in table style (e.g., TableGrid) to the table.
        table.StyleIdentifier = StyleIdentifier.TableGrid;

        // Specify which parts of the style should be applied.
        table.StyleOptions = TableStyleOptions.FirstRow | TableStyleOptions.RowBands;

        // Adjust column widths to fit the content.
        table.AutoFit(AutoFitBehavior.AutoFitToContents);

        // Save the document as a Markdown file.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("TableWithStyle.md", saveOptions);
    }
}
