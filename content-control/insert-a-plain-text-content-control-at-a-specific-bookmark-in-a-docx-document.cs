using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlAtBookmark
{
    public class Program
    {
        public static void Main()
        {
            // Create a sample document with a bookmark.
            Document seedDoc = new Document();
            DocumentBuilder seedBuilder = new DocumentBuilder(seedDoc);
            seedBuilder.Writeln("This is a sample document.");
            seedBuilder.StartBookmark("TargetBookmark");
            seedBuilder.Writeln("Text inside the bookmark.");
            seedBuilder.EndBookmark("TargetBookmark");
            seedDoc.Save("input.docx");

            // Load the document that contains the bookmark.
            Document doc = new Document("input.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Move the cursor to the start of the bookmark.
            builder.MoveToBookmark("TargetBookmark");

            // Insert a plain‑text content control (structured document tag) at the bookmark position.
            StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            sdt.Title = "CustomerName";
            sdt.Tag = "customer-name";

            // Replace any default placeholder with custom text.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "Contoso"));

            // Save the modified document.
            doc.Save("output.docx");
        }
    }
}
