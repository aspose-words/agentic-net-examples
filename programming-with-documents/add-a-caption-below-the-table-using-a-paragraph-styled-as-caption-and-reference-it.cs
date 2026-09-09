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

        // Build a simple 2x2 table.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell 1");
        builder.InsertCell();
        builder.Write("Cell 2");
        builder.EndRow();

        builder.InsertCell();
        builder.Write("Cell 3");
        builder.InsertCell();
        builder.Write("Cell 4");
        builder.EndRow();
        builder.EndTable();

        // Insert a caption paragraph styled as "Caption" and bookmark it for referencing.
        builder.StartBookmark("TableCaption");
        builder.ParagraphFormat.StyleName = "Caption";
        builder.Writeln("Table 1: Sample table.");
        builder.EndBookmark("TableCaption");

        // Move cursor to the end of the document to add a reference to the caption.
        builder.MoveToDocumentEnd();

        // Insert a reference field that points to the bookmarked caption.
        builder.Write("See Table ");
        builder.InsertField(" REF TableCaption \\h ");
        builder.Writeln(" for details.");

        // Save the document to the local file system.
        doc.Save("TableCaptionReference.docx");
    }
}
