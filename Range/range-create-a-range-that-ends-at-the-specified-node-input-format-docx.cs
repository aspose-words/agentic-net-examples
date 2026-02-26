using System;
using Aspose.Words;

class CreateRangeEndingAtNode
{
    static void Main()
    {
        // Load an existing DOCX document.
        Document doc = new Document("InputDocument.docx");

        // Locate the node that will be the end of the range.
        // For this example we use the first paragraph in the body.
        Node endNode = doc.FirstSection.Body.FirstParagraph;

        // The Range property of a node returns a view that starts at the beginning of the
        // document and ends at the end of the specified node. Use the fully‑qualified type
        // name to avoid the ambiguity with System.Range introduced in C# 8.0.
        Aspose.Words.Range rangeEndingAtNode = endNode.Range;

        // Example usage: output the text covered by the range.
        Console.WriteLine("Range text up to the specified node:");
        Console.WriteLine(rangeEndingAtNode.Text.Trim());

        // Optionally, save the document (no changes made in this example).
        doc.Save("OutputDocument.docx");
    }
}
