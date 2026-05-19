using System;
using System.Drawing; // System.Drawing.Color is required for font color
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom character style.
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        customStyle.Font.Name = "Courier New";
        customStyle.Font.Size = 24;
        customStyle.Font.Color = Color.Blue; // Use System.Drawing.Color

        // Create a block‑level rich‑text content control and assign the custom style.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            StyleName = "MyCustomStyle"
        };

        // Add a paragraph with a run inside the content control.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.AppendChild(new Run(doc, "Styled text inside rich text content control."));
        richTextSdt.AppendChild(paragraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the resulting document.
        doc.Save("StyledRichTextContentControl.docx");
    }
}
