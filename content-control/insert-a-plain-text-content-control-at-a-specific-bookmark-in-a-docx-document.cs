using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a bookmark.
        Document seedDoc = new Document();
        DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
        seedBuilder.Writeln("This is a sample paragraph before the bookmark.");
        seedBuilder.StartBookmark("MyBookmark");
        seedBuilder.Writeln("Text inside the bookmark.");
        seedBuilder.EndBookmark("MyBookmark");
        seedBuilder.Writeln("This is a sample paragraph after the bookmark.");

        const string inputPath = "input.docx";
        seedDoc.Save(inputPath);

        // Load the document that contains the bookmark.
        Document doc = new Document(inputPath);

        // Verify that the bookmark exists.
        if (doc.Range.Bookmarks["MyBookmark"] == null)
        {
            throw new InvalidOperationException("Bookmark 'MyBookmark' was not found.");
        }

        // Move the builder to the bookmark location.
        DocumentBuilder docBuilder = new DocumentBuilder(doc);
        docBuilder.MoveToBookmark("MyBookmark");

        // Create a plain text content control (StructuredDocumentTag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "InsertedControl",
            Tag = "inserted-control"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Inserted Text"));

        // Insert the content control at the bookmark position.
        docBuilder.InsertNode(sdt);

        // Save the updated document.
        const string outputPath = "output.docx";
        doc.Save(outputPath);
    }
}
