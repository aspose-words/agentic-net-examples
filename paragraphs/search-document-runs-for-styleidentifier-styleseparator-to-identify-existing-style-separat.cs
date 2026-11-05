using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Write the first part of the paragraph.
        builder.Writeln("First part of the paragraph.");

        // Insert a style separator. This creates a Run that contains the Unicode
        // character U+001F (unit separator) which Word treats as a style separator.
        builder.InsertStyleSeparator();

        // Write the second part of the paragraph.
        builder.Writeln("Second part after the style separator.");

        // Save the document so it can be inspected if needed.
        doc.Save("Output.docx");

        // Search for runs that are style separators.
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        int separatorCount = 0;

        foreach (Run run in runs)
        {
            // A style separator run contains the special character U+001F.
            if (run.Text == "\u001F")
            {
                separatorCount++;

                // Output information about the containing paragraph.
                Paragraph parentParagraph = (Paragraph)run.GetAncestor(NodeType.Paragraph);
                Console.WriteLine($"Found style separator in paragraph: \"{parentParagraph.GetText().Trim()}\"");
            }
        }

        Console.WriteLine($"Total style separators found: {separatorCount}");
    }
}
