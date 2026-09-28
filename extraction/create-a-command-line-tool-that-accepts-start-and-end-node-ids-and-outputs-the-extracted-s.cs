using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class ExtractionTool
{
    public static void Main(string[] args)
    {
        // Expect two arguments: startNodeId endNodeId
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: ExtractionTool <startNodeId> <endNodeId>");
            return;
        }

        string startId = args[0];
        string endId = args[1];

        // Create a sample document if it does not exist
        const string samplePath = "sample.docx";
        if (!File.Exists(samplePath))
        {
            CreateSampleDocument(samplePath);
        }

        // Load the document
        Document sourceDoc = new Document(samplePath);

        // Find start and end paragraphs by their text (used as IDs)
        Paragraph startParagraph = FindParagraphByText(sourceDoc, startId);
        Paragraph endParagraph = FindParagraphByText(sourceDoc, endId);

        if (startParagraph == null)
            throw new InvalidOperationException($"Start node with Id '{startId}' not found.");
        if (endParagraph == null)
            throw new InvalidOperationException($"End node with Id '{endId}' not found.");

        // Ensure start appears before end in document order
        if (!IsNodeBefore(sourceDoc, startParagraph, endParagraph))
            throw new InvalidOperationException("Start node must appear before end node.");

        // Collect nodes between start and end (inclusive)
        List<Node> nodesInRange = CollectNodesInRange(sourceDoc, startParagraph, endParagraph);

        // Build a new document with the extracted nodes
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        foreach (Node node in nodesInRange)
        {
            if (node is Paragraph para)
            {
                body.AppendChild((Paragraph)para.Clone(true));
            }
            else if (node is Table tbl)
            {
                body.AppendChild((Table)tbl.Clone(true));
            }
        }

        // Save the extracted segment as PDF
        const string outputPdf = "extracted.pdf";
        extractedDoc.Save(outputPdf, SaveFormat.Pdf);

        // Validate output
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Failed to create the PDF output.");

        // Confirmation file
        File.WriteAllText("extraction-success.txt", $"Extracted segment saved to {outputPdf}");
    }

    private static void CreateSampleDocument(string path)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Add several paragraphs with identifiable text
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        // Insert a table to demonstrate mixed content
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndRow();
        builder.InsertCell();
        builder.Write("Cell A2");
        builder.InsertCell();
        builder.Write("Cell B2");
        builder.EndRow();
        builder.EndTable();

        // Add more paragraphs after the table
        for (int i = 6; i <= 8; i++)
        {
            builder.Writeln($"Paragraph {i}");
        }

        doc.Save(path);
    }

    private static Paragraph FindParagraphByText(Document doc, string textId)
    {
        foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
        {
            // Trim to ignore trailing paragraph marks
            if (para.GetText().Trim() == textId)
                return para;
        }
        return null;
    }

    private static bool IsNodeBefore(Document doc, Node first, Node second)
    {
        NodeCollection allNodes = doc.GetChildNodes(NodeType.Any, true);
        bool firstSeen = false;
        foreach (Node node in allNodes)
        {
            if (node == first)
                firstSeen = true;
            if (node == second)
                return firstSeen;
        }
        return false;
    }

    private static List<Node> CollectNodesInRange(Document doc, Node startNode, Node endNode)
    {
        List<Node> rangeNodes = new List<Node>();
        bool collecting = false;
        foreach (Node node in doc.GetChildNodes(NodeType.Any, true))
        {
            if (node == startNode)
                collecting = true;

            if (collecting)
                rangeNodes.Add(node);

            if (node == endNode)
                break;
        }
        return rangeNodes;
    }
}
