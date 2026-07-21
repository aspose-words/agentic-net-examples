using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionExample
{
    public static void Main()
    {
        // Create a sample document with four paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");

        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // Load the document we just created.
        Document loadedDoc = new Document(sourcePath);

        // Retrieve all paragraphs in the document.
        NodeCollection paragraphs = loadedDoc.GetChildNodes(NodeType.Paragraph, true);
        if (paragraphs.Count < 4)
            throw new InvalidOperationException("The sample document does not contain the expected paragraphs.");

        // Intentionally set the start node after the end node to trigger error handling.
        Paragraph startParagraph = (Paragraph)paragraphs[2]; // "Paragraph 3"
        Paragraph endParagraph = (Paragraph)paragraphs[1];   // "Paragraph 2"

        try
        {
            // This call should raise an exception because the start node appears after the end node.
            Document extracted = ExtractContentBetweenNodes(loadedDoc, startParagraph, endParagraph);
            extracted.Save("extracted-invalid.docx");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Handled error: {ex.Message}");
        }

        // Correct the order: start before end.
        startParagraph = (Paragraph)paragraphs[1]; // "Paragraph 2"
        endParagraph = (Paragraph)paragraphs[2];   // "Paragraph 3"

        // Perform a valid extraction.
        Document validExtraction = ExtractContentBetweenNodes(loadedDoc, startParagraph, endParagraph);
        const string resultPath = "extracted-valid.docx";
        validExtraction.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not saved correctly.");

        Console.WriteLine("Extraction completed successfully.");
    }

    /// <summary>
    /// Extracts the content between two paragraph nodes (inclusive) into a new document.
    /// Throws an exception if the start node appears after the end node in the source document.
    /// </summary>
    private static Document ExtractContentBetweenNodes(Document source, Paragraph start, Paragraph end)
    {
        // Ensure both nodes belong to the same document.
        if (start.Document != source || end.Document != source)
            throw new InvalidOperationException("Start or end paragraph does not belong to the source document.");

        // Determine the positions of the start and end paragraphs within the source document.
        NodeCollection allParagraphs = source.GetChildNodes(NodeType.Paragraph, true);
        int startIndex = allParagraphs.IndexOf(start);
        int endIndex = allParagraphs.IndexOf(end);

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("One or both of the specified paragraphs were not found in the document.");

        // Validate ordering.
        if (startIndex > endIndex)
            throw new InvalidOperationException("Start paragraph appears after the end paragraph. Extraction aborted.");

        // Create a new empty document to hold the extracted content.
        Document result = new Document();
        result.RemoveAllChildren(); // Ensure the document is empty.

        // Build the minimal required structure: Section -> Body.
        Section section = new Section(result);
        result.AppendChild(section);
        Body body = new Body(result);
        section.AppendChild(body);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

        // Clone and import each paragraph from start to end (inclusive).
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = (Paragraph)allParagraphs[i];
            Node importedNode = importer.ImportNode(srcParagraph, true);
            body.AppendChild(importedNode);
        }

        return result;
    }
}
