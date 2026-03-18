using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Markup;

class ApplyStyleToRichTextContentControl
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // 1. Define a custom character style that will be applied to the SDT.
        // -----------------------------------------------------------------
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = Color.Blue;

        // -----------------------------------------------------------------
        // 2. Create a rich‑text StructuredDocumentTag (content control).
        // -----------------------------------------------------------------
        StructuredDocumentTag richTextTag = new StructuredDocumentTag(
            doc,
            SdtType.RichText,          // Rich‑text content control
            MarkupLevel.Inline);       // Inline level (inside a paragraph)

        // Apply the custom style to the content control.
        richTextTag.StyleName = customStyle.Name;

        // Insert the content control into the document.
        builder.InsertNode(richTextTag);

        // -----------------------------------------------------------------
        // 3. Add text inside the content control – it will inherit the style.
        // -----------------------------------------------------------------
        builder.MoveTo(richTextTag);
        builder.Write("This text is styled by the custom style.");

        // -----------------------------------------------------------------
        // 4. Save the document.
        // -----------------------------------------------------------------
        doc.Save("StyledContentControl.docx");
    }
}
