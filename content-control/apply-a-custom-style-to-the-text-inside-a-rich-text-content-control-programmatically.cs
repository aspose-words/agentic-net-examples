using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Markup;
using Newtonsoft.Json;

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

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        richTextSdt.Title = "RichTextControl";

        // Create a paragraph inside the content control and apply the custom style.
        Paragraph paragraph = new Paragraph(doc);
        // Apply the custom style to the paragraph via ParagraphFormat.
        paragraph.ParagraphFormat.Style = customStyle;
        paragraph.AppendChild(new Run(doc, "This text is inside a rich text content control with a custom style."));

        // Add the paragraph to the content control.
        richTextSdt.AppendChild(paragraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the resulting document.
        doc.Save("styled-richtext-sdt.docx");

        // Optional: write style information to a JSON file using Newtonsoft.Json.
        var styleInfo = new { StyleName = customStyle.Name, Font = customStyle.Font.Name, Size = customStyle.Font.Size };
        File.WriteAllText("style-info.json", JsonConvert.SerializeObject(styleInfo, Formatting.Indented));
    }
}
