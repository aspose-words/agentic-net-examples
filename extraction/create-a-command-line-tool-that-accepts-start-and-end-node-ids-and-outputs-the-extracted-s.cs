using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class Program
{
    // Entry point of the console application.
    // Expects two integer arguments: startNodeId endNodeId
    public static void Main(string[] args)
    {
        // Validate command‑line arguments.
        if (args.Length < 2 ||
            !int.TryParse(args[0], out int startNodeId) ||
            !int.TryParse(args[1], out int endNodeId))
        {
            Console.WriteLine("Usage: dotnet run <startNodeId> <endNodeId>");
            return;
        }

        // -----------------------------------------------------------------
        // 1. Create a sample source document with identifiable nodes.
        // -----------------------------------------------------------------
        const string sourcePath = "sample.docx";

        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Create five paragraphs and assign a deterministic CustomNodeId to each.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
            // The paragraph just created is the last paragraph in the body.
            Paragraph para = sourceDoc.FirstSection.Body.Paragraphs[sourceDoc.FirstSection.Body.Paragraphs.Count - 1];
            para.CustomNodeId = i; // Use the loop index as the node identifier.
        }

        // Persist the sample document to disk.
        sourceDoc.Save(sourcePath);

        // -----------------------------------------------------------------
        // 2. Load the document and locate the start and end nodes by ID.
        // -----------------------------------------------------------------
        Document loadedDoc = new Document(sourcePath);

        NodeCollection allNodes = loadedDoc.GetChildNodes(NodeType.Any, true);
        int startIndex = -1;
        int endIndex = -1;

        for (int i = 0; i < allNodes.Count; i++)
        {
            Node node = allNodes[i];
            if (node.CustomNodeId == startNodeId)
                startIndex = i;
            if (node.CustomNodeId == endNodeId)
                endIndex = i;
        }

        if (startIndex == -1)
            throw new InvalidOperationException($"Start node with CustomNodeId {startNodeId} not found.");
        if (endIndex == -1)
            throw new InvalidOperationException($"End node with CustomNodeId {endNodeId} not found.");
        if (startIndex > endIndex)
            throw new InvalidOperationException("Start node appears after end node in the document order.");

        // -----------------------------------------------------------------
        // 3. Build a new document containing the extracted segment.
        // -----------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Start with an empty document.

        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        for (int i = startIndex; i <= endIndex; i++)
        {
            Node node = allNodes[i];

            // Only block‑level nodes (Paragraph, Table, etc.) can be appended directly to Body.
            // Inline nodes must be wrapped inside a Paragraph.
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
            {
                body.AppendChild(node.Clone(true));
            }
            else
            {
                Paragraph wrapper = new Paragraph(extractedDoc);
                wrapper.AppendChild(node.Clone(true));
                body.AppendChild(wrapper);
            }
        }

        // -----------------------------------------------------------------
        // 4. Save the extracted segment as PDF.
        // -----------------------------------------------------------------
        const string outputPdf = "extracted.pdf";
        extractedDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created.
        if (!File.Exists(outputPdf))
            throw new InvalidOperationException("Failed to create the PDF output file.");

        Console.WriteLine($"Extraction complete. PDF saved to '{outputPdf}'.");
    }
}
