using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add some sample paragraphs – headings and normal text.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("First Heading (Heading 1)");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("This is a normal paragraph.");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Writeln("Second Heading (Heading 2)");

        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Another normal paragraph.");

        // Create a custom paragraph style that will be applied to all headings.
        Style customHeadingStyle = doc.Styles.Add(StyleType.Paragraph, "MyCustomHeading");
        customHeadingStyle.Font.Color = System.Drawing.Color.Red;          // Example formatting.
        customHeadingStyle.Font.Size = 16;
        customHeadingStyle.Font.Bold = true;

        // Loop through all paragraphs in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);
        foreach (Paragraph para in paragraphs)
        {
            // Check if the paragraph uses any built‑in heading style.
            if (para.ParagraphFormat.IsHeading)
            {
                // Apply the custom style to the heading paragraph.
                para.ParagraphFormat.StyleName = customHeadingStyle.Name;
            }
        }

        // Ensure the output directory exists.
        string outputDir = Path.Combine(Environment.CurrentDirectory, "Output");
        Directory.CreateDirectory(outputDir);

        // Save the modified document.
        string outputPath = Path.Combine(outputDir, "StyledHeadings.docx");
        doc.Save(outputPath);

        Console.WriteLine($"Document saved to: {outputPath}");
    }
}
