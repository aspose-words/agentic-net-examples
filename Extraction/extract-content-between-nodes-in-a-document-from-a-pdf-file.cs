using System;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the PDF file as a Word document.
        Document pdfDoc = new Document("input.pdf");

        // Locate the start and end markers using XPath.
        // Adjust the XPath expressions to match the actual markers in your PDF.
        Node startNode = pdfDoc.SelectSingleNode("//Paragraph[contains(., 'START_MARKER')]");
        Node endNode = pdfDoc.SelectSingleNode("//Paragraph[contains(., 'END_MARKER')]");

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end marker not found.");
            return;
        }

        // Create a new empty document to hold the extracted content.
        Document extractedDoc = new Document();

        // Get the body of the new document where nodes will be inserted.
        var targetBody = extractedDoc.FirstSection.Body;

        // Iterate over the nodes that lie between the start and end markers (exclusive).
        Node current = startNode.NextSibling;
        while (current != null && !current.Equals(endNode))
        {
            // Import the node into the new document while preserving formatting.
            Node importedNode = extractedDoc.ImportNode(current, true);
            targetBody.AppendChild(importedNode);
            current = current.NextSibling;
        }

        // Save the extracted portion to a new file.
        extractedDoc.Save("extracted.docx");
    }
}
