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

            // Insert a heading paragraph.
            builder.Writeln("Document with a content control to be duplicated:");
            builder.Writeln();

            // Create a block‑level rich‑text content control.
            StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Title = "SampleControl",
                Tag = "sample"
            };

            // Add a paragraph with some text inside the content control.
            Paragraph innerParagraph = new Paragraph(doc);
            innerParagraph.AppendChild(new Run(doc, "Original content control text."));
            originalSdt.AppendChild(innerParagraph);

            // Append the original content control to the document body.
            doc.FirstSection.Body.AppendChild(originalSdt);

            // Add a blank paragraph after the original control for visual separation.
            doc.FirstSection.Body.AppendChild(new Paragraph(doc));

            // Clone the original content control (deep clone) and insert the copy after the original.
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)originalSdt.Clone(true);
            // Optionally modify the cloned control (e.g., change its title or text).
            clonedSdt.Title = "SampleControlCopy";
            clonedSdt.Tag = "sampleCopy";
            // Change the inner text of the cloned control.
            if (clonedSdt.FirstChild is Paragraph clonedParagraph && clonedParagraph.FirstChild is Run clonedRun)
            {
                clonedRun.Text = "Cloned content control text.";
            }

            // Insert the cloned control after the original one.
            doc.FirstSection.Body.InsertAfter(clonedSdt, originalSdt);

            // Save the resulting document.
            doc.Save("DuplicatedContentControl.docx");
        }
    }
}
