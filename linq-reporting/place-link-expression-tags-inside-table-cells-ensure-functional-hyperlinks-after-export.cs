using System;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables; // <-- added namespace for Table

class Program
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a table with two rows and two columns.
        Table table = builder.StartTable(); // now Table is recognized

        // First row, first cell – insert a hyperlink.
        builder.InsertCell();
        // InsertHyperlink adds a HYPERLINK field. The display text is "Aspose", the URL is the expression.
        // The third argument (false) indicates that the second parameter is a URL, not a bookmark.
        builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);

        // First row, second cell – plain text.
        builder.InsertCell();
        builder.Write("Plain text");

        // End first row.
        builder.EndRow();

        // Second row, first cell – another hyperlink with a different expression.
        builder.InsertCell();
        builder.InsertHyperlink("GitHub", "https://github.com/aspose-words", false);

        // Second row, second cell – plain text.
        builder.InsertCell();
        builder.Write("More text");

        // End the table.
        builder.EndTable();

        // Configure Markdown save options.
        // Use Inline mode so that links are written directly in the Markdown output.
        MarkdownSaveOptions saveOptions = new MarkdownSaveOptions
        {
            LinkExportMode = MarkdownLinkExportMode.Inline
        };

        // Save the document as Markdown.
        doc.Save("TableWithLinks.md", saveOptions);
    }
}
