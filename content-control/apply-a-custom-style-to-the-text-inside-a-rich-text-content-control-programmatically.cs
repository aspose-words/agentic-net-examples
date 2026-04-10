using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Ensure the document has at least one section.
        Section section = doc.FirstSection ?? (Section)doc.AppendChild(new Section(doc));

        // Create a custom character style (character style is required for runs).
        Style customStyle = doc.Styles.Add(StyleType.Character, "MyCustomStyle");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 14;

        // Insert a block‑level rich‑text content control into the body.
        StructuredDocumentTag richTextSdt = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);
        section.Body.AppendChild(richTextSdt);

        // Add a paragraph inside the content control.
        Paragraph paragraph = new Paragraph(doc);
        richTextSdt.AppendChild(paragraph);

        // Add a run with some text.
        Run run = new Run(doc, "This text is inside a rich text content control with a custom style.");
        paragraph.AppendChild(run);

        // Apply the custom character style to the run.
        run.Font.StyleName = "MyCustomStyle";

        // Save the document.
        doc.Save("Output.docx");
    }
}
