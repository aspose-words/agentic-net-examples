using System;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a custom character style named "MyCustomStyle".
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        customStyle.Font.Name = "Calibri";
        customStyle.Font.Size = 14;
        customStyle.Font.Color = System.Drawing.Color.DarkBlue;
        customStyle.Font.Bold = true;

        // Create a rich text content control (block level).
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "RichTextControl",
            Tag = "RichTextTag",
            // Apply the custom style to the content control.
            Style = customStyle
        };

        // Add a paragraph with some text inside the content control.
        Paragraph innerParagraph = new Paragraph(doc);
        Run run = new Run(doc, "This text is inside a rich text content control with a custom style.");
        innerParagraph.AppendChild(run);
        richTextSdt.AppendChild(innerParagraph);

        // Insert the content control into the document body.
        doc.FirstSection.Body.AppendChild(richTextSdt);

        // Save the document to the working directory.
        doc.Save("RichTextContentControlWithCustomStyle.docx");
    }
}
