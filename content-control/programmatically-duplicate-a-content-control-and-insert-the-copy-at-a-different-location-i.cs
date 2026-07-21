using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlDuplication
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add some introductory text.
            builder.Writeln("Document with a content control that will be duplicated.");

            // Create a block‑level rich‑text content control.
            StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "SampleControl",
                Tag = "sample-control"
            };

            // Add a paragraph with some text inside the content control.
            Paragraph innerParagraph = new Paragraph(doc);
            innerParagraph.AppendChild(new Run(doc, "This is the original content control."));
            originalSdt.AppendChild(innerParagraph);

            // Insert the content control into the document body.
            doc.FirstSection.Body.AppendChild(originalSdt);

            // Add a paragraph after the original content control.
            builder.Writeln("Text after the original content control.");

            // Clone the original content control (deep clone, including its children).
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)originalSdt.Clone(true);

            // Optionally modify the cloned control (e.g., change its title/tag or inner text).
            clonedSdt.Title = "ClonedControl";
            clonedSdt.Tag = "cloned-control";

            // Change the inner text of the cloned control.
            if (clonedSdt.FirstChild is Paragraph clonedParagraph && clonedParagraph.FirstChild is Run clonedRun)
            {
                clonedRun.Text = "This is the cloned content control.";
            }

            // Insert the cloned content control after the original one.
            doc.FirstSection.Body.InsertAfter(clonedSdt, originalSdt);

            // Save the resulting document.
            doc.Save("DuplicatedContentControl.docx");
        }
    }
}
