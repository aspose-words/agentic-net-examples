using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;

namespace AsposeWordsExtractBetweenNodes
{
    class Program
    {
        static void Main()
        {
            // Load the HTML file into an Aspose.Words document.
            // The file path can be changed to point to the actual HTML source.
            Document doc = new Document("input.html");

            // Define the XPath expressions for the start and end nodes.
            // Adjust the XPath to match the actual nodes you want to use as boundaries.
            // Example assumes paragraphs with id attributes "start" and "end".
            Node startNode = doc.SelectSingleNode("//p[@id='start']");
            Node endNode = doc.SelectSingleNode("//p[@id='end']");

            if (startNode == null || endNode == null)
            {
                Console.WriteLine("Start or end node not found.");
                return;
            }

            // Collect the text of all nodes that lie between the start and end nodes (exclusive).
            // This loop walks the sibling chain from the node after startNode up to (but not including) endNode.
            string extractedText = string.Empty;
            for (Node curNode = startNode.NextSibling; curNode != null && curNode != endNode; curNode = curNode.NextSibling)
            {
                // Use GetText to retrieve the visible text of each node.
                extractedText += curNode.GetText();
            }

            // Optionally, trim the result to remove leading/trailing whitespace.
            extractedText = extractedText.Trim();

            // Output the extracted content to the console.
            Console.WriteLine("Extracted content between nodes:");
            Console.WriteLine(extractedText);

            // Save the extracted content to a plain text file.
            File.WriteAllText("ExtractedContent.txt", extractedText);

            // If you need the extracted fragment as a Word document, you can create a new document,
            // import the nodes, and save it. Below is an example of saving as a DOCX.
            Document fragmentDoc = new Document();
            // Ensure the new document has at least one section and body.
            fragmentDoc.RemoveAllChildren();
            Section section = new Section(fragmentDoc);
            fragmentDoc.AppendChild(section);
            Body body = new Body(fragmentDoc);
            section.AppendChild(body);

            // Import each node between the boundaries into the new document.
            NodeImporter importer = new NodeImporter(doc, fragmentDoc, ImportFormatMode.KeepSourceFormatting);
            for (Node curNode = startNode.NextSibling; curNode != null && curNode != endNode; curNode = curNode.NextSibling)
            {
                Node importedNode = importer.ImportNode(curNode, true);
                body.AppendChild(importedNode);
            }

            // Save the fragment document.
            fragmentDoc.Save("ExtractedFragment.docx", SaveFormat.Docx);
        }
    }
}
