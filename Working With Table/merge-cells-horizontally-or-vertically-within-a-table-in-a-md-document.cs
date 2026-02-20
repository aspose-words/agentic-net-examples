using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start building a table.
        Table table = builder.StartTable();

        // ---------- First row: horizontal merge ----------
        // First cell – mark as the first cell in a horizontal merge range.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.First;
        builder.Write("Horizontally merged cells");

        // Second cell – merge with the previous cell.
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.Previous;
        // No content needed for the merged cell.
        builder.EndRow();

        // ---------- Second row: vertical merge ----------
        // First cell – mark as the first cell in a vertical merge range.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.First;
        builder.Write("Vertically merged cells");

        // Second cell – merge with the cell above.
        builder.InsertCell();
        builder.CellFormat.VerticalMerge = CellMerge.Previous;
        builder.EndRow();

        // ---------- Third row: regular cells ----------
        builder.InsertCell();
        builder.CellFormat.HorizontalMerge = CellMerge.None;
        builder.CellFormat.VerticalMerge = CellMerge.None;
        builder.Write("Normal cell 1");

        builder.InsertCell();
        builder.Write("Normal cell 2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // Save the document as Markdown.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions();
        doc.Save("MergedTable.md", saveOptions);
    }
}
