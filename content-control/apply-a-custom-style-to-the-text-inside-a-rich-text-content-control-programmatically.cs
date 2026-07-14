using System;
using System.Drawing;
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

        // Add a custom paragraph style to the document.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = Color.Blue;
        customStyle.Font.Bold = true;

        // Create a block‑level rich‑text content control.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(
            doc,
            SdtType.RichText,
            MarkupLevel.Block);
        richTextSdt.Title = "RichTextControl";
        richTextSdt.Tag = "rich-text";

        // Create a paragraph inside the content control and apply the custom style.
        Paragraph para = new Paragraph(doc);
        para.ParagraphFormat.Style = customStyle;

        // Add a run with sample text.
        Run run = new Run(doc, "This text is inside a rich text content control with a custom style.");
        para.AppendChild(run);

        // Assemble the document structure.
        richTextSdt.AppendChild(para);
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the resulting document.
        doc.Save("StyledRichTextSdt.docx");

        // Serialize style information to JSON (demonstrates Newtonsoft.Json usage).
        var styleInfo = new
        {
            StyleName = customStyle.Name,
            FontName = customStyle.Font.Name,
            FontSize = customStyle.Font.Size,
            ColorArgb = customStyle.Font.Color.ToArgb()
        };
        string json = JsonConvert.SerializeObject(styleInfo, Formatting.Indented);
        File.WriteAllText("StyleInfo.json", json);
    }
}
