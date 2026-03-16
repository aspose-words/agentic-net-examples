using System;
using System.Linq;
using Aspose.Words;

class ParagraphKeywordCounter
{
    static void Main()
    {
        // Create or load a document.
        // Here we load an existing document from disk.
        Document doc = new Document("input.docx");

        // The keyword to search for (case‑insensitive).
        string keyword = "Aspose";

        // Get all paragraph nodes in the document.
        NodeCollection paragraphs = doc.GetChildNodes(NodeType.Paragraph, true);

        // Count paragraphs that contain the keyword.
        int count = paragraphs
            .Cast<Paragraph>()
            .Count(p => p.GetText()
                         .IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0);

        // Output the result.
        Console.WriteLine($"Number of paragraphs containing \"{keyword}\": {count}");

        // Save the document (optional – demonstrates the required save rule).
        doc.Save("output.docx");
    }
}
