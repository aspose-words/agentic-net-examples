using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;

public class ApplyStyleToHeadings
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some heading paragraphs using built‑in heading styles.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2");

        // Add a normal paragraph for contrast.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        // Create a custom paragraph style that will be applied to all headings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomHeading");
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 16;
        customStyle.Font.Color = Color.Blue;
        customStyle.ParagraphFormat.SpaceAfter = 12;

        // Loop through all paragraphs in the document and apply the custom style
        // to those that are recognized as headings (built‑in Heading styles).
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.IsHeading)
            {
                para.ParagraphFormat.StyleName = customStyle.Name;
            }
        }

        // Save the resulting document.
        doc.Save("StyledHeadings.docx");
    }
}
