using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Saving;

public class ExtractionTool
{
    public static void Main(string[] args)
    {
        // Parse command‑line arguments (start and end node IDs)
        int startId = 0;
        int endId = 0;
        if (args.Length >= 2)
        {
            int.TryParse(args[0], out startId);
            int.TryParse(args[1], out endId);
        }
        else
        {
            // Default range if arguments are missing
            startId = 1;
            endId = 3;
        }

        // Create a sample source document
        const string sourcePath = "sample.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Paragraph 0");
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.StartTable();
        builder.InsertCell(); builder.Write("Cell A1");
        builder.InsertCell(); builder.Write("Cell B1");
        builder.EndRow();
        builder.InsertCell(); builder.Write("Cell A2");
        builder.InsertCell(); builder.Write("Cell B2");
        builder.EndRow();
        builder.EndTable();
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        sourceDoc.Save(sourcePath);

        // Load the source document
        Document loadedDoc = new Document(sourcePath);
        Body body = loadedDoc.FirstSection.Body;

        // Collect block‑level nodes (Paragraphs and Tables) in document order
        List<Node> blockNodes = new List<Node>();
        foreach (Node node in body.GetChildNodes(NodeType.Any, true))
        {
            if (node.NodeType == NodeType.Paragraph || node.NodeType == NodeType.Table)
            {
                blockNodes.Add(node);
            }
        }

        // Validate indices
        if (startId < 0 || endId < 0 || startId >= blockNodes.Count || endId >= blockNodes.Count || startId > endId)
        {
            throw new InvalidOperationException("Invalid start or end node IDs.");
        }

        // Prepare the result document
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();
        Section resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        Body resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import and append nodes within the specified range
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        for (int i = startId; i <= endId; i++)
        {
            Node srcNode = blockNodes[i];
            Node importedNode = importer.ImportNode(srcNode, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted segment as PDF
        const string outputPdf = "extracted.pdf";
        resultDoc.Save(outputPdf, SaveFormat.Pdf);

        // Verify that the PDF was created
        if (!File.Exists(outputPdf))
        {
            throw new InvalidOperationException("Failed to create the PDF output.");
        }
    }
}
