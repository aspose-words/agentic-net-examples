using System;
using System.Drawing;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add paragraphs with built‑in heading styles and normal text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        // Create a custom paragraph style that will be applied to all headings.
        Style customHeadingStyle = doc.Styles.Add(StyleType.Paragraph, "MyHeadingStyle");
        customHeadingStyle.Font.Color = Color.Blue;
        customHeadingStyle.Font.Size = 14;
        customHeadingStyle.Font.Bold = true;

        // Loop through all paragraphs and apply the custom style to headings.
        NodeCollection allParagraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in allParagraphs)
        {
            if (para.ParagraphFormat.IsHeading)
            {
                para.ParagraphFormat.Style = customHeadingStyle;
            }
        }

        // Save the resulting document.
        doc.Save("StyledHeadings.docx");
    }
}
