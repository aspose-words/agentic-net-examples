using System;
using System.Linq;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Tables;

class RetrieveHeadings
{
    static void Main()
    {
        // Load the DOTX template document.
        Document doc = new Document("input.dotx");

        // Get all paragraph nodes in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Filter paragraphs whose style name starts with "Heading" (e.g., Heading 1, Heading 2, etc.).
        List<string> headings = paragraphs
            .Cast<Paragraph>()
            .Where(p => p.ParagraphFormat.StyleName.StartsWith("Heading", StringComparison.OrdinalIgnoreCase))
            .Select(p => p.GetText().Trim()) // GetText() returns the paragraph text including the paragraph break.
            .ToList();

        // Output the headings.
        foreach (string heading in headings)
        {
            Console.WriteLine(heading);
        }

        // Optionally save the document after processing.
        doc.Save("output.docx");
    }
}
