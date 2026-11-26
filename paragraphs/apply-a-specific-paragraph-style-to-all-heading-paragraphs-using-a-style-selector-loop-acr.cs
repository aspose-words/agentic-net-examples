using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample heading paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Insert a normal paragraph.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a regular paragraph that should remain unchanged.");

        // Insert another heading.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1: Details");

        // Create a custom paragraph style to apply to all headings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyHeadingStyle");
        customStyle.Font.Color = Color.Red;
        customStyle.Font.Bold = true;
        customStyle.Font.Size = 14;

        // Loop through all paragraphs and apply the custom style to heading paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;
            // Check if the paragraph uses any built‑in heading style.
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
                // Apply the custom style.
                para.ParagraphFormat.StyleName = "MyHeadingStyle";
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
