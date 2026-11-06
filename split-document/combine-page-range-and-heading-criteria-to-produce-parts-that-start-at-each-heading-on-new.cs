using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample source document with headings and page breaks.
        string sourcePath = "Source.docx";
        CreateSampleDocument(sourcePath);

        // Load the source document.
        Document sourceDoc = new Document(sourcePath);

        // Collect all Heading 1 paragraphs.
        NodeCollection paragraphs = sourceDoc.GetChildNodes(NodeType.Paragraph, true);
        List<Paragraph> headingNodes = new List<Paragraph>();

        foreach (Paragraph para in paragraphs)
        {
            if (para.ParagraphFormat.StyleIdentifier == StyleIdentifier.Heading1)
                headingNodes.Add(para);
        }

        if (headingNodes.Count == 0)
            throw new InvalidOperationException("No Heading 1 paragraphs found in the source document.");

        // Split the document at each heading.
        for (int i = 0; i < headingNodes.Count; i++)
        {
            Paragraph startHeading = headingNodes[i];
            // Node that marks the end of the current part (the next heading or null for the last part).
            Node endNode = (i + 1 < headingNodes.Count) ? headingNodes[i + 1] : null;

            // Create a new empty document for the part.
            Document partDoc = new Document();
            Body destBody = partDoc.FirstSection.Body;

            // Import nodes from the start heading up to (but not including) the end node.
            Node currentNode = startHeading;
            NodeImporter importer = new NodeImporter(sourceDoc, partDoc, ImportFormatMode.KeepSourceFormatting);

            while (currentNode != null && currentNode != endNode)
            {
                Node importedNode = importer.ImportNode(currentNode, true);
                destBody.AppendChild(importedNode);
                currentNode = currentNode.NextSibling;
            }

            // Save the split part.
            string partPath = $"Part_{i + 1}.docx";
            partDoc.Save(partPath);

            // Verify that the file was created.
            if (!File.Exists(partPath))
                throw new InvalidOperationException($"Failed to create split part file: {partPath}");
        }

        // All parts created successfully.
    }

    // Helper method to create a sample document containing three chapters,
    // each starting with a Heading 1 paragraph and separated by page breaks.
    private static void CreateSampleDocument(string filePath)
    {
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);

        // Chapter 1
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 1");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 1. This paragraph will span multiple lines to increase page count.");
        builder.Writeln("More content in chapter 1.");

        // Page break before Chapter 2
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 2
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 2");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 2.");
        builder.Writeln("Additional content for chapter 2.");

        // Page break before Chapter 3
        builder.InsertBreak(BreakType.PageBreak);

        // Chapter 3
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builder.Writeln("Chapter 3");
        builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Normal;
        builder.Writeln("Content of chapter 3.");
        builder.Writeln("Final content.");

        // Save the sample document.
        doc.Save(filePath);
    }
}
