using System;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add sample content: headings and normal paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph under the first heading.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2: Details");

        // Define the target style to apply to all heading paragraphs.
        // For this example we use the built‑in "Intense Quote" style.
        StyleIdentifier targetStyle = StyleIdentifier.IntenseQuote;

        // Loop through all paragraphs in the document.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Check if the paragraph uses any built‑in heading style.
            StyleIdentifier currentStyle = para.ParagraphFormat.StyleIdentifier;
            if (currentStyle == StyleIdentifier.Heading1 ||
                currentStyle == StyleIdentifier.Heading2 ||
                currentStyle == StyleIdentifier.Heading3 ||
                currentStyle == StyleIdentifier.Heading4 ||
                currentStyle == StyleIdentifier.Heading5 ||
                currentStyle == StyleIdentifier.Heading6 ||
                currentStyle == StyleIdentifier.Heading7 ||
                currentStyle == StyleIdentifier.Heading8 ||
                currentStyle == StyleIdentifier.Heading9)
            {
                // Apply the target style to the heading paragraph.
                para.ParagraphFormat.StyleIdentifier = targetStyle;
            }
        }

        // Save the modified document.
        doc.Save("StyledHeadings.docx");
    }
}
