using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Lists;

public class Program
{
    public static void Main()
    {
        // Create a new empty document.
        Document doc = new Document();

        // Use DocumentBuilder to add sample paragraphs, some with built‑in heading styles.
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Heading 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1: Introduction");

        // Normal paragraph
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph that should remain unchanged.");

        // Heading 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Section 1.1: Overview");

        // Another normal paragraph
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("More regular content.");

        // Heading 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Subsection 1.1.1: Details");

        // Create a custom paragraph style that will be applied to all headings.
        Style customHeadingStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomHeading");
        customHeadingStyle.Font.Name = "Arial";
        customHeadingStyle.Font.Size = 16;
        customHeadingStyle.Font.Color = Color.DarkBlue;
        customHeadingStyle.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        customHeadingStyle.ParagraphFormat.SpaceAfter = 12;

        // Loop through all paragraphs in the document and replace the style
        // of any paragraph that is a heading with the custom style.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // The IsHeading property is true for built‑in heading styles.
            if (para.ParagraphFormat.IsHeading)
            {
                // Apply the custom style.
                para.ParagraphFormat.Style = customHeadingStyle;
            }
        }

        // Save the resulting document.
        doc.Save("Output.docx");
    }
}
