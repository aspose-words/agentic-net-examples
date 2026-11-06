using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

namespace ContentControlStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Add a custom character style that will be applied to the text inside the content control.
            Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 14;
            customStyle.Font.Color = System.Drawing.Color.Blue; // Use a distinct color for demonstration.

            // Create a block-level rich text content control.
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
            richTextSdt.Title = "RichTextControl";
            richTextSdt.Tag = "rich-text";

            // Create a paragraph to hold the text inside the content control.
            Paragraph paragraph = new Paragraph(doc);
            // Create a run with sample text.
            Run run = new Run(doc, "This text is inside a rich text content control and uses a custom style.");
            // Apply the custom style to the run.
            run.Font.StyleName = customStyle.Name;

            // Assemble the nodes.
            paragraph.AppendChild(run);
            richTextSdt.AppendChild(paragraph);
            // Insert the content control into the document body.
            doc.FirstSection.Body.AppendChild(richTextSdt);

            // Save the resulting document.
            doc.Save("StyledRichTextContentControl.docx");
        }
    }
}
