using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Use DocumentBuilder to position the cursor.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Please fill in the following field:");

            // Create an inline plain‑text content control.
            StructuredDocumentTag contentControl = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "CustomerName",          // Friendly name for identification.
                Tag = "customer-name"            // Machine‑readable identifier.
            };

            // Add placeholder text inside the control.
            contentControl.RemoveAllChildren();
            contentControl.AppendChild(new Run(doc, "Enter name here"));

            // Insert the content control into the document.
            builder.InsertNode(contentControl);

            // Save the document.
            const string outputPath = "ContentControlTitleTag.docx";
            doc.Save(outputPath);

            // Reload the document to demonstrate lookup by Title and Tag.
            Document loadedDoc = new Document(outputPath);

            // Find the control by its Title.
            IStructuredDocumentTag? byTitle = loadedDoc.Range.StructuredDocumentTags.GetByTitle("CustomerName");
            if (byTitle != null)
            {
                Console.WriteLine($"Found control by Title: Title='{byTitle.Title}', Tag='{byTitle.Tag}'");
            }

            // Find the control by its Tag.
            IStructuredDocumentTag? byTag = loadedDoc.Range.StructuredDocumentTags.GetByTag("customer-name");
            if (byTag != null)
            {
                Console.WriteLine($"Found control by Tag: Title='{byTag.Title}', Tag='{byTag.Tag}'");
            }
        }
    }
}
