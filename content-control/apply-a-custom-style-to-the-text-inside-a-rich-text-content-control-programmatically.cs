using System;
using System.Drawing; // For font color handling
using Aspose.Words;
using Aspose.Words.Markup;

namespace ContentControlStyleExample
{
    public class Program
    {
        public static void Main()
        {
            // Create a new blank document.
            Document doc = new Document();

            // Define a custom character style.
            Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
            customStyle.Font.Name = "Arial";
            customStyle.Font.Size = 14;
            customStyle.Font.Color = Color.Blue; // System.Drawing.Color

            // Create a block‑level rich‑text content control and assign the custom style.
            StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
            {
                Style = customStyle
            };

            // Add a paragraph with a run of text inside the content control.
            Paragraph innerParagraph = new Paragraph(doc);
            innerParagraph.AppendChild(new Run(doc, "Hello styled content control!"));
            richTextSdt.AppendChild(innerParagraph);

            // Insert the content control into the document body.
            doc.FirstSection.Body.AppendChild(richTextSdt);

            // Save the resulting document.
            doc.Save("StyledRichTextContentControl.docx");
        }
    }
}
