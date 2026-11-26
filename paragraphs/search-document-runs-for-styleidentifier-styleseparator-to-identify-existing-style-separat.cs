using System;
using Aspose.Words;

public class Program
{
    public static void Main()
    {
        // Create a new document and a DocumentBuilder for editing.
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // First part of the paragraph with Arial 12pt.
        builder.Font.Name = "Arial";
        builder.Font.Size = 12;
        builder.Writeln("This is the first part of the paragraph.");

        // Insert a style separator. This creates a Run node that contains the
        // Unicode character U+001F (unit separator) and is hidden.
        builder.InsertStyleSeparator();

        // Second part of the paragraph with Times New Roman 14pt.
        builder.Font.Name = "Times New Roman";
        builder.Font.Size = 14;
        builder.Writeln("This is the second part after the style separator.");

        // Save the document.
        doc.Save("Output.docx");

        // Search for style separator runs.
        NodeCollection runs = doc.GetChildNodes(NodeType.Run, true);
        int separatorCount = 0;
        foreach (Run run in runs)
        {
            // In versions where Run.IsStyleSeparator is unavailable, a style separator
            // is represented by the hidden unit‑separator character (U+001F).
            if (run.Text == "\u001F")
            {
                separatorCount++;
                // Find the paragraph that contains this separator.
                Paragraph parentParagraph = (Paragraph)run.GetAncestor(NodeType.Paragraph);
                Console.WriteLine($"Separator {separatorCount} is in paragraph: \"{parentParagraph.GetText().Trim()}\"");
            }
        }

        Console.WriteLine($"Found {separatorCount} style separator(s) in the document.");
    }
}
