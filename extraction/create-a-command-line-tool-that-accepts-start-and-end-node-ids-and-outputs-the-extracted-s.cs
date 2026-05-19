using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main(string[] args)
    {
        // Validate command‑line arguments.
        if (args.Length < 2)
        {
            Console.Error.WriteLine("Usage: <program> <startNodeId> <endNodeId>");
            return;
        }

        if (!int.TryParse(args[0], out int startId) || !int.TryParse(args[1], out int endId))
        {
            Console.Error.WriteLine("Node IDs must be integer values.");
            return;
        }

        // -----------------------------------------------------------------
        // Create a sample source document with identifiable paragraphs.
        // -----------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Build five paragraphs and assign a custom node identifier to each.
        for (int i = 1; i <= 5; i++)
        {
            builder.Writeln($"Paragraph {i}");
            Paragraph para = (Paragraph)sourceDoc.GetChild(NodeType.Paragraph, sourceDoc.GetChildNodes(NodeType.Paragraph, true).Count - 1, true);
            para.CustomNodeId = i; // Use the loop index as the node ID.
        }

        // -----------------------------------------------------------------
        // Locate the start and end nodes by their CustomNodeId values.
        // -----------------------------------------------------------------
        NodeCollection allParagraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        Paragraph startParagraph = allParagraphs
            .OfType<Paragraph>()
            .FirstOrDefault(p => p.CustomNodeId == startId);
        Paragraph endParagraph = allParagraphs
            .OfType<Paragraph>()
            .FirstOrDefault(p => p.CustomNodeId == endId);

        if (startParagraph == null || endParagraph == null)
        {
            Console.Error.WriteLine("One or both specified node IDs were not found in the document.");
            return;
        }

        int startIndex = allParagraphs.IndexOf(startParagraph);
        int endIndex = allParagraphs.IndexOf(endParagraph);

        if (startIndex > endIndex)
        {
            Console.Error.WriteLine("Start node must appear before or at the same position as the end node.");
            return;
        }

        // -----------------------------------------------------------------
        // Create a new document that will contain the extracted segment.
        // -----------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Ensure a clean document structure.

        // Add a single section with a body.
        Section section = new Section(extractedDoc);
        extractedDoc.AppendChild(section);
        Body body = new Body(extractedDoc);
        section.AppendChild(body);

        // Use NodeImporter to preserve formatting while importing nodes.
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        // Clone and import each paragraph within the requested range.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcPara = (Paragraph)allParagraphs[i];
            Node importedNode = importer.ImportNode(srcPara, true);
            body.AppendChild(importedNode);
        }

        // -----------------------------------------------------------------
        // Save the extracted segment as a PDF file.
        // -----------------------------------------------------------------
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ExtractedSegment.pdf");
        extractedDoc.Save(outputPath, SaveFormat.Pdf);

        // Indicate successful completion.
        Console.WriteLine($"Extracted segment saved to: {outputPath}");
    }
}
