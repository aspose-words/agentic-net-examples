using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class Program
{
    static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a TOC field that will only list entries located inside the bookmark "TblBookmark".
        // \b switch limits the TOC to the bookmark's range.
        // \h makes entries clickable, \z hides page numbers for entries without a page, \u builds the outline.
        builder.InsertTableOfContents("\\b TblBookmark \\h \\z \\u");

        // Add a page break so the TOC appears on its own page.
        builder.InsertBreak(BreakType.PageBreak);

        // Start a bookmark that will surround the table and its heading.
        builder.StartBookmark("TblBookmark");

        // Insert a heading inside the bookmark; this heading will become the TOC entry for the table.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Table 1 – Sample Data");

        // Build a simple 2×2 table.
        Table table = builder.StartTable();

        // Header row.
        builder.InsertCell();
        builder.Write("Column A");
        builder.InsertCell();
        builder.Write("Column B");
        builder.EndRow();

        // Data row.
        builder.InsertCell();
        builder.Write("A1");
        builder.InsertCell();
        builder.Write("B1");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // End the bookmark.
        builder.EndBookmark("TblBookmark");

        // Update all fields (including the TOC) so the entry appears.
        doc.UpdateFields();

        // Save the document.
        doc.Save("TOC_Table_Bookmark.docx");
    }
}
