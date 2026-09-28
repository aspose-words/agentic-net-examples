using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new document and a builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some initial content.
        builder.Writeln("This is a sample document.");

        // Insert a bookmark where the content control will be placed.
        builder.StartBookmark("InsertHere");
        builder.Writeln("Text inside the bookmark.");
        builder.EndBookmark("InsertHere");

        // Save the intermediate document (optional, demonstrates the source file).
        doc.Save("input.docx");

        // Move the builder to the bookmark location.
        builder.MoveToBookmark("InsertHere");

        // Create a plain text content control (inline).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "InsertedControl",
            Tag = "inserted-control"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Hello World"));

        // Insert the content control at the bookmark position.
        builder.InsertNode(sdt);

        // Save the resulting document.
        doc.Save("plain-text-sdt-at-bookmark.docx");
    }
}
