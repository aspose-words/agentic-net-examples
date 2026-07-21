using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to add content.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a paragraph that will host the content control.
            builder.Writeln("Customer information:");

            // Create an inline plain‑text content control (StructuredDocumentTag).
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",   // Friendly name for identification.
                Tag = "customer-name"     // Machine‑readable identifier.
            };

            // Set placeholder text inside the content control.
            sdt.RemoveAllChildren();
            sdt.AppendChild(new Run(doc, "Enter name here"));

            // Append the content control to the last paragraph of the document body.
            Paragraph para = doc.FirstSection.Body.LastParagraph;
            para.AppendChild(sdt);

            // Save the document.
            const string outputPath = "ContentControlTitleTag.docx";
            doc.Save(outputPath);
            Console.WriteLine($"Document saved to {outputPath}");

            // Retrieve the content control by Title.
            // GetByTitle returns IStructuredDocumentTag, so cast to StructuredDocumentTag.
            StructuredDocumentTag? foundByTitle = doc.Range.StructuredDocumentTags.GetByTitle("CustomerName") as StructuredDocumentTag;
            if (foundByTitle != null)
            {
                Console.WriteLine($"Found SDT by Title: Title='{foundByTitle.Title}', Tag='{foundByTitle.Tag}'");
            }

            // Retrieve the content control by Tag.
            StructuredDocumentTag? foundByTag = doc.Range.StructuredDocumentTags.GetByTag("customer-name") as StructuredDocumentTag;
            if (foundByTag != null)
            {
                Console.WriteLine($"Found SDT by Tag: Title='{foundByTag.Title}', Tag='{foundByTag.Tag}'");
            }
        }
    }
}
