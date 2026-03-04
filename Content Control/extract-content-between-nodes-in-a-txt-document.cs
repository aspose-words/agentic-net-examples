using System;
using System.IO;
using System.Text;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source TXT document. Aspose.Words treats each line as a paragraph.
        Document doc = new Document("input.txt");

        // Indices of the start and end paragraphs (0‑based). Adjust as needed.
        int startParagraphIndex = 2; // third paragraph in the document
        int endParagraphIndex   = 5; // sixth paragraph in the document

        // Retrieve the start and end nodes.
        Node startNode = doc.GetChild(NodeType.Paragraph, startParagraphIndex, true);
        Node endNode   = doc.GetChild(NodeType.Paragraph, endParagraphIndex, true);

        // Collect text that lies strictly between the two nodes.
        StringBuilder builder = new StringBuilder();

        // Walk forward from the node after startNode up to (but not including) endNode.
        for (Node cur = startNode.NextPreOrder(doc);
             cur != null && cur != endNode;
             cur = cur.NextPreOrder(doc))
        {
            // Append the textual representation of nodes that contain visible text.
            // For a plain‑text document only Paragraph and Run nodes contain text.
            if (cur is Paragraph || cur is Run)
                builder.Append(cur.GetText());
        }

        // The extracted string, trimmed of leading/trailing whitespace.
        string extractedText = builder.ToString().Trim();

        // Display the result.
        Console.WriteLine(extractedText);

        // Optionally, save the extracted portion to a new TXT file.
        File.WriteAllText("extracted.txt", extractedText);
    }
}
