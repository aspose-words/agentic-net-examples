using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample source document.
        const string sourcePath = "source.docx";
        var sourceDoc = new Document();
        var builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 0 - before start");
        builder.Writeln("Paragraph 1 - start");
        builder.Writeln("Paragraph 2 - middle");
        builder.Writeln("Paragraph 3 - end");
        builder.Writeln("Paragraph 4 - after end");
        sourceDoc.Save(sourcePath);

        // Load the document.
        var doc = new Document(sourcePath);
        var body = doc.FirstSection.Body;

        // Identify the start and end paragraphs (by index).
        if (body.Paragraphs.Count < 4)
            throw new InvalidOperationException("The source document does not contain enough paragraphs.");

        var startParagraph = body.Paragraphs[1];
        var endParagraph = body.Paragraphs[3];

        // Determine their positions.
        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || startIndex > endIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // Prepare the result document.
        var resultDoc = new Document();
        resultDoc.RemoveAllChildren(); // Remove the default empty section.

        var resultSection = new Section(resultDoc);
        resultDoc.AppendChild(resultSection);
        var resultBody = new Body(resultDoc);
        resultSection.AppendChild(resultBody);

        // Importer to copy nodes between documents.
        var importer = new NodeImporter(doc, resultDoc, ImportFormatMode.KeepSourceFormatting);

        // Clone and copy paragraphs from start to end inclusive.
        for (int i = startIndex; i <= endIndex; i++)
        {
            var paragraph = body.Paragraphs[i];
            var importedNode = importer.ImportNode(paragraph, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted document.
        const string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Verify that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Optional: write a small JSON report.
        var report = new
        {
            SourceDocument = sourcePath,
            ExtractedDocument = resultPath,
            ExtractedParagraphCount = resultBody.Paragraphs.Count,
            ExtractionTime = DateTime.UtcNow
        };
        string json = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("extractionInfo.json", json);
    }
}
