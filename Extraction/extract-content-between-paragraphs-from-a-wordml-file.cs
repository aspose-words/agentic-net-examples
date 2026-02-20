using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

class ExtractBetweenParagraphs
{
    static void Main()
    {
        // Path to the WORDML (XML) document.
        const string inputPath = @"C:\Docs\input.xml";

        // Load the WORDML document. Aspose.Words automatically detects the format.
        Document doc = new Document(inputPath);

        // Indices of the start and end paragraphs (zero‑based).
        // Adjust these values to the paragraphs that bound the region you want to extract.
        const int startParagraphIndex = 2; // third paragraph in the document
        const int endParagraphIndex   = 5; // sixth paragraph in the document

        // Retrieve the start and end Paragraph nodes.
        Paragraph startParagraph = (Paragraph)doc.GetChild(NodeType.Paragraph, startParagraphIndex, true);
        Paragraph endParagraph   = (Paragraph)doc.GetChild(NodeType.Paragraph, endParagraphIndex, true);

        if (startParagraph == null || endParagraph == null)
        {
            Console.WriteLine("One of the specified paragraph indices is out of range.");
            return;
        }

        // Build a string containing all content that lies strictly between the two paragraphs.
        StringBuilder betweenText = new StringBuilder();

        // Begin with the node immediately after the start paragraph.
        Node currentNode = startParagraph.NextSibling;

        // Walk the node tree until we reach the end paragraph.
        while (currentNode != null && currentNode != endParagraph)
        {
            // Append the textual representation of the current node.
            // GetText() returns the text of the node and all its descendants.
            betweenText.Append(currentNode.GetText());

            // Move to the next node in the document order.
            currentNode = currentNode.NextSibling;
        }

        // Output the extracted text.
        Console.WriteLine("Content between paragraph {0} and paragraph {1}:", startParagraphIndex, endParagraphIndex);
        Console.WriteLine(betweenText.ToString());
    }
}
