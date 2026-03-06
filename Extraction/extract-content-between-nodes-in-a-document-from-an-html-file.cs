using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Saving;

class Program
{
    static void Main()
    {
        // Load the HTML file into an Aspose.Words document.
        Document doc = new Document("input.html");

        // Locate the start and end markers. In this example we look for <div id="start"> and <div id="end">.
        Node startNode = doc.SelectSingleNode("//div[@id='start']");
        Node endNode = doc.SelectSingleNode("//div[@id='end']");

        if (startNode == null || endNode == null)
        {
            Console.WriteLine("Start or end node not found.");
            return;
        }

        // Verify that both nodes share the same parent (they are siblings).
        if (startNode.ParentNode != endNode.ParentNode)
        {
            Console.WriteLine("Start and end nodes are not siblings.");
            return;
        }

        // Collect the HTML representation of all nodes that lie between the start and end nodes.
        StringBuilder sb = new StringBuilder();
        Node current = startNode.NextSibling;
        while (current != null && current != endNode)
        {
            // Export each node to raw HTML using the ToString overload.
            sb.Append(current.ToString(SaveFormat.Html));
            current = current.NextSibling;
        }

        string extractedHtml = sb.ToString();

        // Save the extracted fragment to a separate HTML file.
        System.IO.File.WriteAllText("extracted.html", extractedHtml);

        Console.WriteLine("Content extracted successfully.");
    }
}
