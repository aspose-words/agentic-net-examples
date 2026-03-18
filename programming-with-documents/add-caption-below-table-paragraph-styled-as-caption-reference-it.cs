using System;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;

class AddTableCaptionExample
{
    static void Main()
    {
        // Create a new document and a DocumentBuilder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Build a simple 2x2 table.
        // -----------------------------------------------------------------
        Table table = builder.StartTable();

        // First row.
        builder.InsertCell();
        builder.Write("Cell 1,1");
        builder.InsertCell();
        builder.Write("Cell 1,2");
        builder.EndRow();

        // Second row.
        builder.InsertCell();
        builder.Write("Cell 2,1");
        builder.InsertCell();
        builder.Write("Cell 2,2");
        builder.EndRow();

        // Finish the table.
        builder.EndTable();

        // -----------------------------------------------------------------
        // Insert a caption paragraph below the table.
        // The paragraph is given the built‑in "Caption" style.
        // -----------------------------------------------------------------
        // Start a bookmark that will be used for referencing the caption.
        builder.StartBookmark("TableCaption");

        // Insert a new paragraph and apply the "Caption" style.
        Paragraph captionParagraph = builder.InsertParagraph();
        captionParagraph.ParagraphFormat.StyleName = "Caption";

        // Write the caption text.
        builder.Writeln("Table 1: Sample table created with Aspose.Words.");

        // End the bookmark.
        builder.EndBookmark("TableCaption");

        // -----------------------------------------------------------------
        // Insert a reference to the caption elsewhere in the document.
        // -----------------------------------------------------------------
        builder.Writeln(); // Add a blank line before the reference.
        builder.Write("See ");
        // Insert a REF field that points to the bookmark "TableCaption".
        // The \h switch makes the reference a hyperlink.
        // Use the overload InsertField(string fieldName, string fieldText) which is available in all supported versions.
        builder.InsertField("REF", "TableCaption \\h");
        builder.Writeln(" for more details.");

        // -----------------------------------------------------------------
        // Save the document.
        // -----------------------------------------------------------------
        doc.Save("TableWithCaption.docx");
    }
}
