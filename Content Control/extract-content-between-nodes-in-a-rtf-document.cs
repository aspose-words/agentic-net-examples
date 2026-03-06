using System;
using Aspose.Words;
using Aspose.Words.Markup;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the RTF document.
        Document doc = new Document(@"C:\Docs\Sample.rtf");

        // Find the first ranged Structured Document Tag (SDT) start node.
        // This node marks the beginning of the region we want to extract.
        StructuredDocumentTagRangeStart startTag = doc.GetChild(
            NodeType.StructuredDocumentTagRangeStart, 0, true) as StructuredDocumentTagRangeStart;

        // Find the corresponding end tag. It is linked by the same Id.
        StructuredDocumentTagRangeEnd endTag = doc.GetChild(
            NodeType.StructuredDocumentTagRangeEnd, 0, true) as StructuredDocumentTagRangeEnd;

        if (startTag == null || endTag == null)
        {
            Console.WriteLine("No ranged Structured Document Tag found in the document.");
            return;
        }

        // The Range of the start tag spans all nodes up to (but not including) the end tag.
        // Retrieve the text contained within this range.
        string extractedText = startTag.Range.Text;

        // Output the extracted content.
        Console.WriteLine("Extracted text between the nodes:");
        Console.WriteLine(extractedText.Trim());
    }
}
