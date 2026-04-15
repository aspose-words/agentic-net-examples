using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Insert sample paragraphs with various heading styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;
        builder.Writeln("Heading 3");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Create a custom paragraph style that will be applied to all headings.
        Style customHeadingStyle = doc.Styles.Add(StyleType.Paragraph, "MyHeadingStyle");
        customHeadingStyle.Font.Name = "Arial";
        customHeadingStyle.Font.Size = 14;
        customHeadingStyle.Font.Color = Color.Blue;
        customHeadingStyle.Font.Bold = true;

        // Loop through all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // If the paragraph uses a built‑in heading style, replace it with the custom style.
            if (para.ParagraphFormat.IsHeading)
            {
                para.ParagraphFormat.Style = customHeadingStyle;
            }
        }

        // Save the resulting document.
        doc.Save("StyledHeadings.docx");
    }
}
