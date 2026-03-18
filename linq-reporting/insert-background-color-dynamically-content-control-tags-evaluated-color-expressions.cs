using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Markup;
using Aspose.Words.Tables; // for Paragraph objects

class DynamicBackgroundColorExample
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a block‑level plain‑text content control (SDT).
        // The Tag property holds a simple color expression in the form "bgcolor:#RRGGBB".
        StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.PlainText, MarkupLevel.Block)
        {
            Title = "DynamicColorTag",
            Tag = "bgcolor:#FFCC00" // Example expression – light orange background.
        };
        builder.InsertNode(sdt);
        builder.Writeln(); // Ensure the SDT has a paragraph to contain text.

        // Optional: set default font for the content inside the control.
        sdt.ContentsFont.Name = "Arial";
        sdt.ContentsFont.Size = 12;
        sdt.ContentsFont.Color = Color.Black;

        // Iterate over all content controls in the document, evaluate the color expression,
        // and apply the background shading to every paragraph that belongs to the control.
        foreach (StructuredDocumentTag tag in doc.GetChildNodes(NodeType.StructuredDocumentTag, true))
        {
            // Expected format: "bgcolor:#RRGGBB"
            Match match = Regex.Match(tag.Tag ?? string.Empty,
                @"bgcolor\s*:\s*#(?<hex>[0-9A-Fa-f]{6})",
                RegexOptions.IgnoreCase);

            if (match.Success)
            {
                string hex = match.Groups["hex"].Value;
                Color background = ColorTranslator.FromHtml("#" + hex);

                // Apply the background color to each paragraph inside the content control.
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
