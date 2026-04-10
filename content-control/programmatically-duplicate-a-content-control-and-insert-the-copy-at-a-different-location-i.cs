using System;
using Aspose.Words;
using Aspose.Words.Markup;

namespace DuplicatingContentControl
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Add an introductory paragraph.
            builder.Writeln("Paragraph before the original content control.");

            // Create a plain‑text content control (inline level).
            StructuredDocumentTag originalSdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
            {
                Title = "SampleControl",
                Tag = "SampleTag"
            };

            // Insert the content control into the current paragraph.
            builder.InsertNode(originalSdt);

            // Add some text inside the content control.
            Run innerRun = new Run(doc, "Original content");
            originalSdt.AppendChild(innerRun);

            // Finish the paragraph that contains the original control.
            builder.Writeln();

            // Add another paragraph to separate the original and the cloned control.
            builder.Writeln("Paragraph between the original and the cloned content control.");

            // Clone the original content control (deep clone, including its children).
            StructuredDocumentTag clonedSdt = (StructuredDocumentTag)originalSdt.Clone(true);

            // Insert the cloned content control at the current cursor position.
            builder.InsertNode(clonedSdt);

            // Add a final paragraph after the cloned control.
            builder.Writeln();
            builder.Writeln("Paragraph after the cloned content control.");

            // Save the resulting document to the working directory.
            doc.Save("DuplicatedContentControl.docx");
        }
    }
}
