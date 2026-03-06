using System;
using Aspose.Words;

class CreateRangeEndingAtNode
{
    static void Main()
    {
        // Load the input DOCX document.
        Document doc = new Document("input.docx");

        // Find the node at which the range should end.
        // Here we take the first paragraph in the document as an example.
        Node endNode = doc.GetChild(NodeType.Paragraph, 0, true);

        // Create a range that ends at the specified node.
        // Use the fully‑qualified Aspose.Words.Range type to avoid the conflict with System.Range.
        Aspose.Words.Range rangeEndingAtNode = endNode.Range;

        // Example usage: output the text covered by the range.
        Console.WriteLine("Range text ending at the node:");
        Console.WriteLine(rangeEndingAtNode.Text.Trim());

        // Save the document (unchanged) to the output path.
        doc.Save("output.docx");
    }
}
