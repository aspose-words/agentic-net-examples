using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class MergeTableCellsInMarkdown
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // -------------------------------------------------
        // First row – demonstrate horizontal merging.
        // -------------------------------------------------
        // First cell: start of a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged across two cells");

        // Second cell: merges with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;

        // End the first row.
        builder.EndRow();

        // -------------------------------------------------
        // Second row – start of a vertical merge.
        // -------------------------------------------------
        // First cell: start of a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged across two rows");

        // Second cell: normal (no vertical merge).
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell");

        // End the second row.
        builder.EndRow();

        // -------------------------------------------------
        // Third row – continuation of the vertical merge.
        // -------------------------------------------------
        // First cell: merges with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;

        // Second cell: another normal cell.
        builder.InsertCell();
        builder.Write("Another normal cell");

        // End the third row and the table.
        builder.EndRow();
        builder.EndTable();

        // -------------------------------------------------
        // Save the document as Markdown.
        // -------------------------------------------------
        // Use raw HTML for tables so that merged cells are preserved,
        // because pure Markdown cannot represent merged cells.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            ExportAsHtml = MarkdownExportAsHtml.Tables
        };
        doc.Save("MergedTable.md", saveOptions);
    }
}
