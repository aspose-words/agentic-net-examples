using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // Create a sample source document with several paragraphs.
        string sourcePath = "source.docx";
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);
        builder.Writeln("Paragraph 1");
        builder.Writeln("Paragraph 2 - start");
        builder.Writeln("Paragraph 3 - middle");
        builder.Writeln("Paragraph 4 - end");
        builder.Writeln("Paragraph 5");
        sourceDoc.Save(sourcePath);

        // Load the source document.
        Document loadedDoc = new Document(sourcePath);
        Body body = loadedDoc.FirstSection.Body;

        // Identify the start and end paragraphs (inclusive).
        Paragraph startParagraph = body.Paragraphs[1]; // "Paragraph 2 - start"
        Paragraph endParagraph = body.Paragraphs[3];   // "Paragraph 4 - end"

        if (startParagraph == null || endParagraph == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        int startIndex = body.Paragraphs.IndexOf(startParagraph);
        int endIndex = body.Paragraphs.IndexOf(endParagraph);
        if (startIndex < 0 || endIndex < 0 || endIndex < startIndex)
            throw new InvalidOperationException("Invalid paragraph boundaries.");

        // Prepare the result document.
        Document resultDoc = new Document();
        resultDoc.RemoveAllChildren();

        Section resultSection = (Section)resultDoc.AppendChild(new Section(resultDoc));
        Body resultBody = (Body)resultSection.AppendChild(new Body(resultDoc));

        // Import and append each paragraph within the range.
        NodeImporter importer = new NodeImporter(loadedDoc, resultDoc, ImportFormatMode.KeepSourceFormatting);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph para = body.Paragraphs[i];
            Node importedNode = importer.ImportNode(para, true);
            resultBody.AppendChild(importedNode);
        }

        // Save the extracted content as a new DOCX file.
        string resultPath = "extracted.docx";
        resultDoc.Save(resultPath);

        // Validate that the output file was created.
        if (!File.Exists(resultPath))
            throw new InvalidOperationException("The extracted document was not created.");

        // Write a simple JSON report confirming success.
        var report = new
        {
            SourceDocument = Path.GetFullPath(sourcePath),
            ExtractedDocument = Path.GetFullPath(resultPath),
            ExtractedParagraphCount = resultBody.Paragraphs.Count
        };
        string jsonReport = JsonConvert.SerializeObject(report, Formatting.Indented);
        File.WriteAllText("extraction_report.json", jsonReport);
    }
}
