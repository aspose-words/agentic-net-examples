using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.Saving;

class ExtractAndConvert
{
    static void Main()
    {
        // Load the source document.
        Document srcDoc = new Document("Input.docx");

        // Find the first Run node that will serve as the start point.
        // (Adjust the condition as needed to locate the desired run.)
        Run startRun = null;
        foreach (Run run in srcDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text.Contains("StartMarker")) // example marker text
            {
                startRun = run;
                break;
            }
        }

        if (startRun == null)
            throw new InvalidOperationException("Start run not found.");

        // Collect all nodes that appear after the start run up to (but not including) the next table.
        List<Node> extractedNodes = new List<Node>();
        Node curNode = startRun.NextPreOrder(srcDoc);
        while (curNode != null && curNode.NodeType != NodeType.Table)
        {
            // Store the node for later import.
            extractedNodes.Add(curNode);
            curNode = curNode.NextPreOrder(srcDoc);
        }

        // If no table was found after the run, the loop ends when curNode is null.
        // At this point 'extractedNodes' contains the desired range.

        // Create a new blank document that will hold the extracted content.
        Document destDoc = new Document();
        // Remove the default nodes (section, body, paragraph) to start with a clean structure.
        destDoc.RemoveAllChildren();

        // Build the minimal required structure: Section -> Body.
        Section section = new Section(destDoc);
        destDoc.AppendChild(section);
        Body body = new Body(destDoc);
        section.AppendChild(body);

        // Prepare a NodeImporter to copy nodes from the source to the destination document.
        NodeImporter importer = new NodeImporter(srcDoc, destDoc, ImportFormatMode.KeepSourceFormatting);

        // Import each collected node and add it to the destination body.
        foreach (Node node in extractedNodes)
        {
            Node importedNode = importer.ImportNode(node, true);
            body.AppendChild(importedNode);
        }

        // Save the extracted portion as XPS using XpsSaveOptions.
        XpsSaveOptions xpsOptions = new XpsSaveOptions(); // defaults to SaveFormat.Xps
        destDoc.Save("ExtractedContent.xps", xpsOptions);
    }
}
