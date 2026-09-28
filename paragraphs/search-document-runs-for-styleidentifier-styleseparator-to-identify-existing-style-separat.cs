using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new blank document.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Build a paragraph that contains a style separator.
        builder.Writeln("Paragraph with style separator:");

        // Apply the first paragraph style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Write("Heading1 text");

        // Insert the style separator run.
        builder.InsertStyleSeparator();

        // Apply the second paragraph style.
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;
        builder.Write("Heading2 text");
        builder.Writeln();

        // Save the document (optional, just to visualize the result).
        doc.Save("Output.docx");

        // Search all runs for the style separator character (Unicode 0x0014).
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        int separatorCount = 0;

        foreach (Run run in runs)
        {
            // A style separator is represented by the special character '\u0014'.
            if (run.Text == "\u0014")
            {
                separatorCount++;

                // Get the containing paragraph and output its text.
                Paragraph paragraph = (Paragraph)run.GetAncestor(NodeType.Paragraph);
                Console.WriteLine($"Style separator found in paragraph: \"{paragraph.GetText().Trim()}\"");
            }
        }

        Console.WriteLine($"Total style separators found: {separatorCount}");
    }
}
