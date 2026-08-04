using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // -----------------------------------------------------------------
        // Create a custom paragraph style that will be applied to TOC entries.
        // -----------------------------------------------------------------
        Style tocCustomStyle = doc.Styles.Add(StyleType.Paragraph, "MyTocStyle");
        tocCustomStyle.Font.Name = "Arial";
        tocCustomStyle.Font.Size = 14;
        tocCustomStyle.Font.Color = System.Drawing.Color.DarkBlue;
        tocCustomStyle.ParagraphFormat.SpaceAfter = 6;

        // ---------------------------------------------------------------
        // Insert a Table of Contents field. The switches pick up headings 1‑3.
        // ---------------------------------------------------------------
        builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");
        builder.InsertBreak(BreakType.PageBreak);

        // ---------------------------------------------------------------
        // Add some headings that will appear in the TOC.
        // ---------------------------------------------------------------
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.Writeln("Section 1.2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 2.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 2.1.1");

        // ---------------------------------------------------------------
        // Update all fields so the TOC is populated.
        // ---------------------------------------------------------------
        doc.UpdateFields();

        // ---------------------------------------------------------------
        // Apply the custom style to all TOC entry paragraphs (TOC1‑TOC9).
        // ---------------------------------------------------------------
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier id = para.ParagraphFormat.StyleIdentifier;
            if (id >= StyleIdentifier.Toc1 && id <= StyleIdentifier.Toc9)
            {
                para.ParagraphFormat.Style = tocCustomStyle;
            }
        }

        // ---------------------------------------------------------------
        // Save the document to the current directory.
        // ---------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomTocStyle.docx");
        doc.Save(outputPath);
    }
}
