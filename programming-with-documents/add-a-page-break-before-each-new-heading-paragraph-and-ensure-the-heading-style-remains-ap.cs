using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content: normal paragraph, heading, normal, heading, etc.
        builder.Writeln("This is an introductory paragraph.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Details of section 1.1.");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");

        // Insert a page break before each heading paragraph.
        // Get all paragraph nodes.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        foreach (Paragraph para in paragraphs)
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            if (styleId == StyleIdentifier.Heading1 ||
                styleId == StyleIdentifier.Heading2 ||
                styleId == StyleIdentifier.Heading3 ||
                styleId == StyleIdentifier.Heading4 ||
                styleId == StyleIdentifier.Heading5 ||
                styleId == StyleIdentifier.Heading6 ||
                styleId == StyleIdentifier.Heading7 ||
                styleId == StyleIdentifier.Heading8 ||
                styleId == StyleIdentifier.Heading9)
            {
                // Move the builder to the start of the heading paragraph and insert a page break.
                DocumentBuilder headingBuilder = new DocumentBuilder(doc);
                headingBuilder.MoveTo(para);
                headingBuilder.InsertBreak(BreakType.PageBreak);
                // The heading style remains unchanged.
            }
        }

        // Save the document.
        string outputPath = "Output.docx";
        doc.Save(outputPath);
    }
}
