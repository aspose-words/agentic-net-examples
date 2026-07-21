using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System.Drawing;

public class ApplyCustomTocStyle
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert a Table of Contents field at the beginning of the document.
        // The switches configure the TOC to include heading levels 1‑3 and to create hyperlinks.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add some headings that will be picked up by the TOC.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");

        // Update all fields so the TOC is populated.
        doc.UpdateFields();

        // Create a custom paragraph style that will be applied to TOC entries.
        Style customTocStyle = doc.Styles.Add(StyleType.Paragraph, "MyTocStyle");
        customTocStyle.Font.Name = "Arial";
        customTocStyle.Font.Size = 12;
        customTocStyle.Font.Color = Color.DarkBlue;
        customTocStyle.Font.Bold = true;

        // Apply the custom style to all TOC entry paragraphs (styles TOC1 … TOC9).
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            Style paraStyle = para.ParagraphFormat.Style;
            if (paraStyle != null &&
                paraStyle.StyleIdentifier >= StyleIdentifier.Toc1 &&
                paraStyle.StyleIdentifier <= StyleIdentifier.Toc9)
            {
                para.ParagraphFormat.Style = customTocStyle;
            }
        }

        // Determine an output path in the current working directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTocStyle.docx");

        // Save the document.
        doc.Save(outputPath);
    }
}
