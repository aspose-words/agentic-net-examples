using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create a sample document with several paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2");
        builder.Writeln("Paragraph 3");
        builder.Writeln("Paragraph 4");
        const string sourcePath = "sample.docx";
        sourceDoc.Save(sourcePath);

        // Load the document for extraction.
        Document loadedDoc = new Document(sourcePath);
        Body body = loadedDoc.FirstSection.Body;

        // Intentionally select start node after end node to trigger error handling.
        Paragraph startParagraph = body.Paragraphs[2]; // "Paragraph 3"
        Paragraph endParagraph = body.Paragraphs[1];   // "Paragraph 2"

        try
        {
            // Attempt extraction with invalid node order.
            ExtractRange(loadedDoc, startParagraph, endParagraph, "extracted-invalid.docx");
        }
        catch (InvalidOperationException ex)
        {
            Console.WriteLine($"Error during extraction: {ex.Message}");
        }

        // Now perform a valid extraction where start precedes end.
        startParagraph = body.Paragraphs[1]; // "Paragraph 2"
        endParagraph = body.Paragraphs[2];   // "Paragraph 3"

        ExtractRange(loadedDoc, startParagraph, endParagraph, "extracted-valid.docx");

        // Verify that the valid extraction output was created.
        if (!File.Exists("extracted-valid.docx"))
            throw new InvalidOperationException("Expected extraction output was not created.");
    }

    private static void ExtractRange(Document source, Paragraph start, Paragraph end, string outputPath)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (start == null) throw new ArgumentNullException(nameof(start));
        if (end == null) throw new ArgumentNullException(nameof(end));
        if (string.IsNullOrEmpty(outputPath)) throw new ArgumentException("Output path must be provided.", nameof(outputPath));

        // Determine the parent Body that contains both paragraphs.
        Body parentBody = start.ParentNode as Body ?? end.ParentNode as Body;
        if (parentBody == null)
            throw new InvalidOperationException("Start or end paragraph does not belong to a Body node.");

        int startIndex = parentBody.Paragraphs.IndexOf(start);
        int endIndex = parentBody.Paragraphs.IndexOf(end);

        if (startIndex == -1 || endIndex == -1)
            throw new InvalidOperationException("Start or end paragraph not found in the document body.");

        // Validate ordering: start must come before or be the same as end.
        if (startIndex > endIndex)
            throw new InvalidOperationException("Start node appears after the end node. Extraction aborted.");

        // Create a new empty document to hold the extracted content.
        Document result = new Document();
        result.RemoveAllChildren();

        // Build the minimal required structure: Section -> Body.
        Section section = new Section(result);
        result.AppendChild(section);
        Body resultBody = new Body(result);
        section.AppendChild(resultBody);

        // Use NodeImporter to import nodes from the source document into the result document.
        NodeImporter importer = new NodeImporter(source, result, ImportFormatMode.KeepSourceFormatting);

        // Clone and import each paragraph from start to end inclusive.
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = parentBody.Paragraphs[i];
            Paragraph importedPara = importer.ImportNode(para, true) as Paragraph;
            if (importedPara != null)
                resultBody.AppendChild(importedPara);
        }

        // Save the extracted content.
        result.Save(outputPath);
    }
}
