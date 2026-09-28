using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a sample source document containing a run marker,
        //    a paragraph between the marker and a table.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph.");
        // This run will be the start marker.
        builder.Write("RunStart");
        builder.Writeln(); // End the paragraph containing the run.

        builder.Writeln("Paragraph between.");

        // Insert a table after the paragraphs.
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell1");
        builder.EndRow();
        builder.EndTable();

        // Optional: save the source document for reference.
        sourceDoc.Save("source.docx");

        // ------------------------------------------------------------
        // 2. Locate the run node that marks the start of the extraction.
        // ------------------------------------------------------------
        Run targetRun = null;
        foreach (Run run in sourceDoc.GetChildNodes(NodeType.Run, true))
        {
            if (run.Text == "RunStart")
            {
                targetRun = run;
                break;
            }
        }

        if (targetRun == null)
            throw new InvalidOperationException("Target run not found.");

        // The paragraph that contains the run.
        Paragraph startParagraph = targetRun.ParentNode as Paragraph;
        if (startParagraph == null)
            throw new InvalidOperationException("Run does not have a parent paragraph.");

        // ------------------------------------------------------------
        // 3. Find the next table after the start paragraph.
        // ------------------------------------------------------------
        Node currentNode = startParagraph.NextSibling;
        Table followingTable = null;
        while (currentNode != null)
        {
            if (currentNode.NodeType == NodeType.Table)
            {
                followingTable = (Table)currentNode;
                break;
            }
            currentNode = currentNode.NextSibling;
        }

        if (followingTable == null)
            throw new InvalidOperationException("Following table not found.");

        // ------------------------------------------------------------
        // 4. Collect all block-level nodes from the start paragraph up
        //    to (but not including) the table.
        // ------------------------------------------------------------
        List<Node> nodesToExtract = new List<Node>();
        Node nodeIter = startParagraph;
        while (nodeIter != null && nodeIter != followingTable)
        {
            // Only Paragraphs and Tables are valid children of Body.
            if (nodeIter.NodeType == NodeType.Paragraph || nodeIter.NodeType == NodeType.Table)
                nodesToExtract.Add(nodeIter);
            nodeIter = nodeIter.NextSibling;
        }

        // ------------------------------------------------------------
        // 5. Build a new document and import the collected nodes.
        // ------------------------------------------------------------
        Document resultDoc = new Document();
        // Remove the default empty section/body.
        resultDoc.RemoveAllChildren();

        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);

        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Use NodeImporter to bring nodes from sourceDoc into resultDoc.
        NodeImporter importer = new NodeImporter(sourceDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        foreach (Node node in nodesToExtract)
        {
            Node importedNode = importer.ImportNode(node, true);
            resultBody.AppendChild(importedNode);
        }

        // ------------------------------------------------------------
        // 6. Save the extracted portion as XPS.
        // ------------------------------------------------------------
        string outputPath = "extracted.xps";
        resultDoc.Save(outputPath, SaveFormat.Xps);

        // Validate that the XPS file was created.
        if (!File.Exists(outputPath))
            throw new InvalidOperationException("The XPS output file was not created.");
    }
}
