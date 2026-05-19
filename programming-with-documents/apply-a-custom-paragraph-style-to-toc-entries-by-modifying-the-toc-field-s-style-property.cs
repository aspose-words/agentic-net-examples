using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new document and builder.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Create a custom paragraph style for TOC entries.
        Style myTocStyle = doc.Styles.Add(StyleType.Paragraph, "MyTocStyle");
        myTocStyle.Font.Name = "Arial";
        myTocStyle.Font.Size = 12;
        myTocStyle.Font.Color = Color.DarkBlue;
        myTocStyle.ParagraphFormat.SpaceAfter = 6;

        // Insert a TOC field. It will be populated after headings are added.
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // Add headings that the TOC will reference.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // Update fields so the TOC is generated.
        doc.UpdateFields();

        // Apply the custom style to all TOC entry paragraphs (TOC1‑TOC9).
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier id = para.ParagraphFormat.StyleIdentifier;
            if (id >= StyleIdentifier.Toc1 && id <= StyleIdentifier.Toc9)
            {
                para.ParagraphFormat.Style = myTocStyle;
            }
        }

        // Save the document.
        doc.Save("MyToc.docx");
    }
}
