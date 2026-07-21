using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using Aspose.Words.Tables;

public class Program
{
    // Entry point of the console application.
    // Expects two integer arguments: startNodeId and endNodeId.
    public static void Main(string[] args)
    {
        // Validate command‑line arguments.
        if (args.Length < 2 ||
            !int.TryParse(args[0], out int startId) ||
            !int.TryParse(args[1], out int endId))
        {
            Console.WriteLine("Usage: dotnet run <startNodeId> <endNodeId>");
            return;
        }

        // -----------------------------------------------------------------
        // 1. Create a sample source document with identifiable nodes.
        // -----------------------------------------------------------------
        const string sourcePath = "source.docx";
        CreateSampleDocument(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the source document.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document(sourcePath);

        // -----------------------------------------------------------------
        // 3. Locate the start and end nodes by their CustomNodeId.
        // -----------------------------------------------------------------
        Node startNode = FindNodeByCustomId(sourceDoc, startId);
        Node endNode   = FindNodeByCustomId(sourceDoc, endId);

        if (startNode == null || endNode == null)
            throw new InvalidOperationException("One or both node IDs were not found in the document.");

        // Ensure the start node appears before the end node.
        if (IsNodeAfter(startNode, endNode))
        {
            Node temp = startNode;
            startNode = endNode;
            endNode   = temp;
        }

        // -----------------------------------------------------------------
        // 4. Extract the range of nodes between startNode and endNode (inclusive).
        // -----------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build a minimal document structure: Section -> Body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Use NodeImporter to correctly import nodes from the source document.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        Node current = startNode;
        while (current != null)
        {
            // Import the node into the destination document.
            Node imported = importer.ImportNode(current, true);
            body.AppendChild(imported);

            if (current == endNode)
                break;

            current = current.NextSibling;
        }

        // -----------------------------------------------------------------
        // 5. Save the extracted segment as PDF.
        // -----------------------------------------------------------------
        const string outputPdf = "extracted.pdf";
        extractedDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        Console.WriteLine($"Extraction complete. PDF saved to '{outputPdf}'.");
    }

    // Creates a sample DOCX file with several paragraphs.
    // Each paragraph is assigned a unique CustomNodeId for later lookup.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
            // The paragraph just written is the current paragraph of the builder.
            builder.CurrentParagraph.CustomNodeId = i;
        }

        doc.Save(filePath);
        if (!File.Exists(filePath))
            throw new InvalidOperationException("Failed to create the sample source document.");
    }

    // Searches the document for a node whose CustomNodeId matches the supplied id.
    private static Node FindNodeByCustomId(Document doc, int id)
    {
        NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
        foreach (Node node in nodes)
        {
            if (node.CustomNodeId == id)
                return node;
        }
        return null;
    }

    // Determines whether nodeA appears after nodeB in the document order.
    private static bool IsNodeAfter(Node nodeA, Node nodeB)
    {
        Node current = nodeA;
        while (current != null)
        {
            if (current == nodeB)
                return true;
            current = current.PreviousSibling;
        }
        return false;
    }
}
