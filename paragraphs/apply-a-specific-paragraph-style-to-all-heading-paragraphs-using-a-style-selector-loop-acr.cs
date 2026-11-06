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

        // Insert sample heading paragraphs.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("First Heading (Heading 1)");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Second Heading (Heading 2)");

        // Insert a normal paragraph for contrast.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        // Create a custom style that will be applied to all headings.
        Style customStyle = doc.Styles.Add(StyleType.Paragraph, "MyHeadingStyle");
        customStyle.Font.Color = Color.Blue;
        customStyle.Font.Size = 14;
        customStyle.Font.Bold = true;

        // Loop through all paragraphs and apply the custom style to heading paragraphs.
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            StyleIdentifier styleId = para.ParagraphFormat.StyleIdentifier;

            // Check if the paragraph uses any built‑in heading style.
            if (styleId >= StyleIdentifier.Heading1 && styleId <= StyleIdentifier.Heading9)
            {
                // Apply the custom style.
                para.ParagraphFormat.StyleName = "MyHeadingStyle";
            }
        }

        // Save the modified document.
        doc.Save("Output.docx");
    }
}
