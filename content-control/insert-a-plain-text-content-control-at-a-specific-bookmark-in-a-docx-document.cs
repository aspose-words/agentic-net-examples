using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add a paragraph before the bookmark.
            builder.Writeln("Paragraph before bookmark.");

            // Insert a bookmark named "TargetBookmark".
            builder.StartBookmark("TargetBookmark");
            builder.Writeln("This text will be replaced by a content control.");
            builder.EndBookmark("TargetBookmark");

            // Add a paragraph after the bookmark.
            builder.Writeln("Paragraph after bookmark.");

            // Move the cursor to the start of the bookmark.
            builder.MoveToBookmark("TargetBookmark");

            // Insert a plain text content control at the bookmark location.
            StructuredDocumentTag sdt = builder.InsertStructuredDocumentTag(SdtType.PlainText);
            sdt.Title = "CustomerName";
            sdt.Tag = "customer-name";

            // Replace any default placeholder with custom text.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "Contoso"));

            // Save the document.
            doc.Save("output.docx");
        }
    }
}
