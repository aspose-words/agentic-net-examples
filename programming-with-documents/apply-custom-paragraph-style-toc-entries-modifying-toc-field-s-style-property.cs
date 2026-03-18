using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ApplyCustomTocStyle
{
    static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert some headings that the TOC will pick up.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        // Insert a TOC field at the beginning of the document.
        builder.MoveToDocumentStart();
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

        // Update fields so the TOC is populated.
        doc.UpdateFields();

        // Create a custom paragraph style that will be applied to TOC entries.
        Style customTocStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomTocStyle");
        customTocStyle.Font.Name = "Arial";
        customTocStyle.Font.Size = 12;
        customTocStyle.Font.Color = System.Drawing.Color.DarkBlue;
        customTocStyle.ParagraphFormat.SpaceAfter = 6;
        customTocStyle.ParagraphFormat.Alignment = ParagraphAlignment.Left;

        // Apply the custom style to all TOC entry paragraphs (Toc1‑Toc9).
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier id = para.ParagraphFormat.Style?.StyleIdentifier ?? StyleIdentifier.Normal;
            if (id >= StyleIdentifier.Toc1 && id <= StyleIdentifier.Toc9)
                para.ParagraphFormat.Style = customTocStyle;
        }

        // Save the document.
        doc.Save("CustomTocStyle.docx");
    }
}
