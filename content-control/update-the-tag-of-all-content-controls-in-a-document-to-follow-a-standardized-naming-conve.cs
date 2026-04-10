using System;
using Aspose.Words;
using Aspose.Words.Markup;

public class UpdateContentControlTags
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a plain‑text content control.
        StructuredDocumentTag plainTag = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Inline)
        {
            Title = "PlainTextControl",
            Tag = "OldTag1"
        };
        builder.InsertNode(plainTag);
        builder.Writeln("Sample text inside plain‑text control.");

        // Insert a rich‑text content control.
        StructuredDocumentTag richTag = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block)
        {
            Title = "RichTextControl",
            Tag = "OldTag2"
        };
        doc.FirstSection.Body.AppendChild(richTag);
        Paragraph richParagraph = new Paragraph(doc);
        richParagraph.AppendChild(new Run(doc, "Sample text inside rich‑text control."));
        richTag.AppendChild(richParagraph);

        // Insert a checkbox content control.
        StructuredDocumentTag checkTag = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline)
        {
            Title = "CheckboxControl",
            Tag = "OldTag3",
            Checked = true
        };
        builder.InsertNode(checkTag);
        builder.Writeln("Checkbox control.");

        // -----------------------------------------------------------------
        // Update the Tag property of every content control to follow a
        // standardized naming convention: "CC_{sequentialNumber}".
        // -----------------------------------------------------------------
        NodeCollection sdtNodes = doc.GetChildNodes(NodeType.StructuredDocumentTag, true);
        int index = 1;
        foreach (Node node in sdtNodes)
        {
            if (node is StructuredDocumentTag sdt)
            {
                sdt.Tag = $"CC_{index}";
                index++;
            }
        }

        // Save the modified document.
        doc.Save("UpdatedContentControls.docx");
    }
}
