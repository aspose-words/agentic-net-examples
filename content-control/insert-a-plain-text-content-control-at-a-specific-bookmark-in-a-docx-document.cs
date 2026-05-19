using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a sample document that contains a bookmark.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Document with a bookmark:");
        builder.StartBookmark("TargetBookmark");
        builder.Writeln("This text is inside the bookmark.");
        builder.EndBookmark("TargetBookmark");
        sourceDoc.Save("input.docx");

        // Load the document that contains the bookmark.
        Document doc = new Document("input.docx");
        DocumentBuilder docBuilder = new DocumentBuilder(doc);

        // Move the cursor to the start of the bookmark.
        docBuilder.MoveToBookmark("TargetBookmark");

        // Insert a plain‑text content control (StructuredDocumentTag) at the bookmark position.
        StructuredDocumentTag sdt = docBuilder.InsertStructuredDocumentTag(SdtType.PlainText);
        sdt.Title = "CustomerName";
        sdt.Tag = "customer-name";

        // Replace any default placeholder with the desired text.
        sdt.RemoveAllChildren();
        sdt.AppendChild(new Run(doc, "Contoso"));

        // Save the modified document.
        doc.Save("output.docx");
    }
}
