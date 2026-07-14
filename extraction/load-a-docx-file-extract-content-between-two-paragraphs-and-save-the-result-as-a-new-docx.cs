using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main(string[] args)
    {
        // Create a sample source document with several paragraphs.
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1 - Intro");
        builder.Writeln("Paragraph 2 - Start");
        builder.Writeln("Paragraph 3 - Middle");
        builder.Writeln("Paragraph 4 - End");
        builder.Writeln("Paragraph 5 - After");
        const string sourcePath = "source.docx";
        sourceDoc.Save(sourcePath);

        // Load the source document.
        var loadedDoc = new Document(sourcePath);
        var body = loadedDoc.FirstSection?.Body;
        if (body == null)
            throw new InvalidOperationException("The source document does not contain a body.");

        // Identify the start and end paragraphs (by index).
        // Here we choose the second and fourth paragraphs (zero‑based indices 1 and 3).
        if (body.Paragraphs.Count <= 3)
            throw new InvalidOperationException("The source document does not contain enough paragraphs.");

        Paragraph startParagraph = body.Paragraphs[1];
        Paragraph endParagraph = body.Paragraphs[3];

        int startIdx = body.Paragraphs.IndexOf(startParagraph);
        int endIdx = body.Paragraphs.IndexOf(endParagraph);

        if (startIdx < 0 || endIdx < 0 || startIdx > endIdx)
            throw new InvalidOperationException("Invalid paragraph boundaries for extraction.");

        // Prepare the result document with a clean structure.
        var resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove the default section/body.

        var resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        var resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Import and append the selected paragraphs into the result document.
        for (int i = startIdx; i <= endIdx; i++)
        {
            Node importedNode = resultDoc.ImportNode(body.Paragraphs[i], true);
            resultBody.AppendChild(importedNode);
        }

        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Write simple extraction metadata as JSON.
        var metadata = new
        {
            SourceDocument = sourcePath,
            ExtractedDocument = resultPath,
            StartParagraphIndex = startIdx,
            EndParagraphIndex = endIdx,
            ExtractedParagraphCount = endIdx - startIdx + 1
        };
        File.WriteAllText("extraction-metadata.json",
            JsonConvert.SerializeObject(metadata, Formatting.Indented));

        // Validate that the output files were created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        if (!File.Exists("extraction-metadata.json"))
            throw new InvalidOperationException("The extraction metadata file was not created.");
    }
}
