using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Start a bookmark that will bound the TOC entries.
        builder.StartBookmark("TableBookmark");

        // Insert a heading that will appear in the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Table 1: Sample Table");

        // Build a simple 2x2 table.
        Table table = builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndTable();

        // End the bookmark.
        builder.EndBookmark("TableBookmark");

        // Move the cursor to the start of the document to insert the TOC.
        builder.MoveToDocumentStart();

        // Insert a TOC that includes only headings within the bookmark.
        // \b specifies the bookmark, \o "1-1" limits to heading level 1, \h adds hyperlinks.
        builder.InsertTableOfContents("\\b TableBookmark \\o \"1-1\" \\h \\z \\u");

        // Update fields so the TOC reflects the current content.
        doc.UpdateFields();

        // Save the resulting document.
        const string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
