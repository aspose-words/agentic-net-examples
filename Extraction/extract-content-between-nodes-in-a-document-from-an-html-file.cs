using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the HTML file into an Aspose.Words Document.
        Document doc = new Document("input.html");

        // Define XPath expressions for the start and end nodes.
        // Adjust these expressions to match the actual nodes you want to use.
        string startXPath = "//Paragraph[@StyleIdentifier='Heading 1'][1]";
        string endXPath   = "//Paragraph[@StyleIdentifier='Heading 2'][1]";

        // Locate the start and end nodes.
        Node startNode = doc.SelectSingleNode(startXPath);
        Node endNode   = doc.SelectSingleNode(endXPath);

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Ensure both nodes share the same parent; otherwise, extraction logic would need to be more complex.
        if (startNode.ParentNode != endNode.ParentNode)
        {
            Console.WriteLine("Start and end nodes do not share the same parent.");
            return;
        }

        // Extract the text that lies between the start and end nodes (exclusive).
        StringBuilder extractedText = new StringBuilder();
        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
        {
            extractedText.Append(cur.GetText());
        }

        // Output the extracted content.
        Console.WriteLine("Extracted content between the specified nodes:");
        Console.WriteLine(extractedText.ToString());

        // Optionally, save the extracted content to a new document.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Ensure the document is empty.
        // Create a new paragraph and add the extracted text.
        Paragraph para = new Paragraph(extractedDoc);
        Run run = new Run(extractedDoc, extractedText.ToString());
        para.AppendChild(run);
        extractedDoc.FirstSection.Body.AppendChild(para);

        // Save the new document (e.g., as a plain text file).
        extractedDoc.Save("extracted.txt", SaveFormat.Text);
    }
}
