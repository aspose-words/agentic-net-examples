using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document with a bookmark named "TargetBookmark".
        Document seedDoc = new Document();
        DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
        seedBuilder.StartBookmark("TargetBookmark");
        seedBuilder.Write("Initial text inside the bookmark.");
        seedBuilder.EndBookmark("TargetBookmark");
        seedDoc.Save("input.docx");

        // Load the document that contains the bookmark.
        Document doc = new Document("input.docx");

        // Ensure the bookmark exists.
        Bookmark bookmark = doc.Range.Bookmarks["TargetBookmark"];
        if (bookmark == null)
            throw new InvalidOperationException("Bookmark 'TargetBookmark' not found.");

        // Clear the bookmark's existing text.
        bookmark.Text = string.Empty;

        // Move the builder to the bookmark location.
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.MoveToBookmark("TargetBookmark");

        // Create a plain text content control (StructuredDocumentTag).
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "CustomerName",
            Tag = "customer-name"
        };
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Insert the content control at the bookmark position.
        builder.InsertNode(sdt);

        // Save the modified document.
        doc.Save("output.docx");
    }
}
