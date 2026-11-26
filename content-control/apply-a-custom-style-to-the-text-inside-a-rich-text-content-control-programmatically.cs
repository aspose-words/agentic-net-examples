using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Define a custom paragraph style.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = Color.Blue; // System.Drawing.Color is required by Font.Color

        // Create a block-level rich text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextSdt.Title = "RichTextWithCustomStyle";

        // Create a paragraph, apply the custom style, and add some text.
        Paragraph paragraph = new Paragraph(doc);
        paragraph.ParagraphFormat.Style = customStyle;
        paragraph.AppendChild(new Run(doc, "This text is inside a rich text content control with a custom style."));

        // Add the paragraph to the content control.
        richTextSdt.AppendChild(paragraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the resulting document.
        doc.Save("custom-style-sdt.docx");
    }
}
