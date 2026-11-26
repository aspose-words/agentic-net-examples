using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionTest
{
    public static void Main()
    {
        // Create a source document with styled paragraphs.
        Document source = new Document();
        DocumentBuilder builder = new DocumentBuilder(source);

        builder.Writeln("Intro paragraph."); // Paragraph 0

        builder.Font.Bold = true;
        builder.Writeln("Bold paragraph."); // Paragraph 1 (styled)

        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Italic paragraph."); // Paragraph 2 (styled)

        builder.Font.Italic = false;
        builder.Writeln("End paragraph."); // Paragraph 3

        // Identify the start and end paragraphs for extraction.
        Paragraph startPara = source.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = source.FirstSection.Body.Paragraphs[2];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Create a new document to hold the extracted content.
        Document extracted = new Document();
        extracted.RemoveAllChildren();

        Section section = new Section(extracted);
        extracted.AppendChild(section);
        Body body = new Body(extracted);
        section.AppendChild(body);

        // Import the paragraphs preserving formatting.
        NodeImporter importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);
        Node importedStart = importer.ImportNode(startPara, true);
        Node importedEnd = importer.ImportNode(endPara, true);

        body.AppendChild(importedStart);
        body.AppendChild(importedEnd);

        // Validate that the extracted paragraphs retain original styling.
        if (extracted.FirstSection.Body.Paragraphs.Count != 2)
            throw new InvalidOperationException("Extracted paragraph count mismatch.");

        Paragraph clonedStart = extracted.FirstSection.Body.Paragraphs[0];
        Paragraph clonedEnd = extracted.FirstSection.Body.Paragraphs[1];

        if (clonedStart.Runs.Count == 0 || clonedEnd.Runs.Count == 0)
            throw new InvalidOperationException("Extracted paragraphs contain no runs.");

        bool startBold = clonedStart.Runs[0].Font.Bold;
        bool endItalic = clonedEnd.Runs[0].Font.Italic;

        if (!startBold || !endItalic)
            throw new InvalidOperationException("Extracted paragraph styling does not match the source.");

        // Save the extracted document and a JSON report.
        string extractedPath = "extracted.docx";
        extracted.Save(extractedPath);

        var report = new
        {
            ExtractedParagraphCount = extracted.FirstSection.Body.Paragraphs.Count,
            StartParagraphBold = startBold,
            EndParagraphItalic = endItalic,
            Success = true
        };
        string reportPath = "extraction-report.json";
        File.WriteAllText(reportPath, JsonConvert.SerializeObject(report, Formatting.Indented));

        // Indicate success.
        Console.WriteLine("Extraction test passed.");
    }
}
