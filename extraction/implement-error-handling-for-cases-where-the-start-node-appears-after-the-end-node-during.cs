using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with three paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        sourceDoc.Save("input.docx");

        // Load the document for extraction.
        Document loadedDoc = new Document("input.docx");
        Body body = loadedDoc.FirstSection.Body;

        // Attempt extraction where start node appears after end node (should raise error).
        Paragraph startAfterEnd = body.Paragraphs[2]; // Third paragraph
        Paragraph endBeforeStart = body.Paragraphs[1]; // Second paragraph

        try
        {
            ExtractRange(loadedDoc, startAfterEnd, endBeforeStart, "invalid-extraction.docx");
        }
        catch (InvalidOperationException ex)
        {
            // Expected error handling.
            Console.WriteLine($"Error: {ex.Message}");
        }

        // Correct extraction where start node precedes end node.
        Paragraph validStart = body.Paragraphs[0]; // First paragraph
        Paragraph validEnd = body.Paragraphs[1];   // Second paragraph

        ExtractRange(loadedDoc, validStart, validEnd, "extracted.docx");

        // Validate that the extraction output file was created.
        if (!File.Exists("extracted.docx"))
            throw new InvalidOperationException("Expected extraction output was not created.");
    }

    private static void ExtractRange(Document source, Paragraph startNode, Paragraph endNode, string outputPath)
    {
        // Ensure both nodes are not null.
        if (startNode == null || endNode == null)
            throw new InvalidOperationException("Start or end node is null.");

        // Ensure both nodes belong to the same body.
        Body body = startNode.ParentNode as Body;
        if (body == null || endNode.ParentNode != body)
            throw new InvalidOperationException("Start and end nodes must be within the same body.");

        // Determine the positions of the nodes within the body.
        int startIndex = body.Paragraphs.IndexOf(startNode);
        int endIndex = body.Paragraphs.IndexOf(endNode);

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("Start or end node not found in the body.");

        // Validate ordering.
        if (startIndex > endIndex)
            throw new InvalidOperationException("Start node appears after the end node. Extraction aborted.");

        // Create a new document to hold the extracted content.
        Document result = new Document();
        result.RemoveAllChildren();

        Section section = new Section(result);
        result.AppendChild(section);
        Body resultBody = new Body(result);
        section.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

        // Clone and import paragraphs from start to end inclusive.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];
            Node importedNode = importer.ImportNode(para, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted content.
        result.Save(outputPath);
    }
}
