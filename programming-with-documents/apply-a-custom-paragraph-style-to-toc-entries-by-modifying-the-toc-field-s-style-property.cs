using System;
using Aspose.Words;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom paragraph style that will be applied to TOC entries.
        Style customTocStyle = doc.Styles.Add(StyleType.Paragraph, "MyTocStyle");
        customTocStyle.Font.Size = 14;
        customTocStyle.Font.Color = Color.DarkBlue;
        customTocStyle.Font.Bold = true;

        // Insert a Table of Contents that will pick up headings level 1‑3.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings so the TOC has entries.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update all fields so the TOC is generated.
        doc.UpdateFields();

        // Apply the custom style to all TOC entry paragraphs (styles TOC1‑TOC9).
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier id = para.ParagraphFormat.StyleIdentifier;
            if (id >= StyleIdentifier.Toc1 && id <= StyleIdentifier.Toc9)
            {
                para.ParagraphFormat.Style = customTocStyle;
            }
        }

        // Save the document.
        string outputPath = "CustomTocStyle.docx";
        doc.Save(outputPath);
    }
}
