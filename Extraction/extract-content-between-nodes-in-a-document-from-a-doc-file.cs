using System;
using System.Text;
using Aspose.Words;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the source DOC file.
        Document srcDoc = new Document("Input.doc");

        // Define the start and end nodes (example: first and third paragraphs).
        Node startNode = srcDoc.SelectSingleNode("//Paragraph[1]");
        Node endNode   = srcDoc.SelectSingleNode("//Paragraph[3]");

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Collect the text of all nodes from startNode up to and including endNode.
        StringBuilder extractedBuilder = new StringBuilder();
        for (Node cur = startNode; cur != null; cur = cur.NextPreOrder(srcDoc))
        {
            extractedBuilder.Append(cur.GetText());
            if (cur == endNode)
                break;
        }

        string extractedText = extractedBuilder.ToString();

        // Create a new document to hold the extracted content.
        Document resultDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(resultDoc);

        // Insert the extracted text into the new document.
        builder.Writeln(extractedText);

        // Save the result document.
        resultDoc.Save("ExtractedContent.docx");

        Console.WriteLine("Extraction complete. Saved to ExtractedContent.docx");
    }
}
