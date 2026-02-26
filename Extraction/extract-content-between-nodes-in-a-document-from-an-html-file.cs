using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractBetweenNodes
{
    static void Main()
    {
        // Load the HTML file into an Aspose.Words document.
        Document doc = new Document("input.html");

        // Locate the start and end nodes using XPath.
        // Adjust the XPath expressions to match the actual markers in your HTML.
        Node startNode = doc.SelectSingleNode("//p[@id='start']");
        Node endNode   = doc.SelectSingleNode("//p[@id='end']");

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Collect the content of all nodes that appear between the start and end nodes (exclusive).
        StringBuilder extracted = new StringBuilder();

        for (Node cur = startNode.NextSibling; cur != null && cur != endNode; cur = cur.NextSibling)
        {
            // Export each node to its HTML representation.
            extracted.Append(cur.ToString(SaveFormat.Html));
        }

        string extractedHtml = extracted.ToString();

        // Output the extracted HTML to the console.
        Console.WriteLine("Extracted HTML content:");
        Console.WriteLine(extractedHtml);

        // Save the extracted content as a separate HTML file.
        Document tempDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(tempDoc);
        builder.InsertHtml(extractedHtml);
        tempDoc.Save("extracted.html", SaveFormat.Html);
    }
}
