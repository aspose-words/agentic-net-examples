using System;
using System.IO;
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

        // Add some sample paragraphs with different built‑in styles.
        builder.Writeln("This is a normal paragraph."); // Normal style

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - First");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Heading 2 - First");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Heading 1 - Second");

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
            // If the paragraph is a built‑in heading, replace its style with the custom one.
            if (para.ParagraphFormat.IsHeading)
            {
                para.ParagraphFormat.Style = customHeadingStyle;
            }
        }

        // Save the resulting document to the current directory.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "StyledHeadings.docx");
        doc.Save(outputPath);
    }
}
