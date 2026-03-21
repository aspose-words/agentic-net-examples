using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables;

class DynamicBackgroundColorExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();

        // Create a block‑level plain‑text content control (SDT) with a tag that holds a color expression.
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "DynamicColorTag",
            Tag = "bgcolor:#FFCC00" // Example expression – light orange background.
        };

        // The SDT must contain at least one block‑level node (e.g., a paragraph) to be valid.
        Paragraph sdtParagraph = new Paragraph(doc);
        sdt.AppendChild(sdtParagraph);
        sdtParagraph.AppendChild(new Run(doc, "Sample text inside the content control."));

        // Insert the SDT into the document body.
        doc.FirstSection.Body.AppendChild(sdt);

        // Optional: set default font for the content inside the control.
        sdt.ContentsFont.Name = "Arial";
        sdt.ContentsFont.Size = 12;
        sdt.ContentsFont.Color = Color.Black;

        // Iterate over all content controls, evaluate the color expression,
        // and apply the background shading to every paragraph that belongs to the control.
        foreach (StructuredDocumentTag tag in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            Match match = Regex.Match(tag.Tag ?? string.Empty,
                @"bgcolor\s*:\s*#(?<hex>[0-9A-Fa-f]{6})",
                RegexOptions.IgnoreCase);

            if (match.Success)
            {
                string hex = match.Groups["hex"].Value;
                Color background = ColorTranslator.FromHtml("#" + hex);

                foreach (Paragraph para in tag.GetChildNodes(NodeType.Paragraph, true))
                {
                    para.ParagraphFormat.Shading.BackgroundPatternColor = background;
                }
            }
        }

        // Save the resulting document.
        doc.Save("DynamicBackgroundColor.docx");
    }
}
