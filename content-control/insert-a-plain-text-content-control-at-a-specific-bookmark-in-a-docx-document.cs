using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a bookmark.
        Document seedDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(seedDoc);
        builder.Writeln("Paragraph before the bookmark.");
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("Placeholder text inside bookmark.");
        builder.EndBookmark("TargetBookmark");
        builder.Writeln("Paragraph after the bookmark.");
        const string inputPath = "sample.docx";
        seedDoc.Save(inputPath);

        // Load the document that contains the bookmark.
        Document doc = new Document(inputPath);

        // Locate the bookmark by name.
        Bookmark bookmark = doc.Range.Bookmarks["TargetBookmark"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark not found.");

        // The bookmark spans nodes inside a paragraph.
        Paragraph parentParagraph = (Paragraph)bookmark.BookmarkStart.ParentNode;

        // Create a plain‑text content control (structured document tag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "InsertedPlainText",
            Tag = "inserted-plain-text"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Inserted content control text."));

        // Insert the content control before the bookmark start node.
        parentParagraph.InsertBefore(sdt, bookmark.BookmarkStart);

        // Remove all nodes that belong to the bookmark (start, end, and any content between them).
        Node current = bookmark.BookmarkStart;
        while (current != null)
        {
            Node next = current.NextSibling;
            parentParagraph.RemoveChild(current);
            if (current == bookmark.BookmarkEnd)
                break;
            current = next;
        }

        // Save the modified document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
