using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class Program
{
    public static void Main()
    {
        // ------------------------------------------------------------
        // 1. Create a source document with styled paragraphs.
        // ------------------------------------------------------------
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        builder.Writeln("Intro paragraph."); // Paragraph 0 (no special style)

        builder.Font.Bold = true;
        builder.Writeln("Styled paragraph one."); // Paragraph 1 (Bold)

        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Styled paragraph two."); // Paragraph 2 (Italic)

        builder.Font.Italic = false;
        builder.Writeln("Ending paragraph."); // Paragraph 3 (no special style)

        // ------------------------------------------------------------
        // 2. Identify the start and end paragraphs for extraction.
        // ------------------------------------------------------------
        Paragraph startPara = sourceDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = sourceDoc.FirstSection.Body.Paragraphs[2];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // ------------------------------------------------------------
        // 3. Prepare a new document that will hold the extracted content.
        // ------------------------------------------------------------
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren(); // Remove the default empty section.

        Section newSection = new Section(extractedDoc);
        extractedDoc.AppendChild(newSection);
        Body newBody = new Body(extractedDoc);
        newSection.AppendChild(newBody);

        // ------------------------------------------------------------
        // 4. Import (clone) the selected paragraphs preserving formatting.
        // ------------------------------------------------------------
        NodeImporter importer = new NodeImporter(sourceDoc, extractedDoc, ImportFormatMode.KeepSourceFormatting);

        Paragraph importedStart = importer.ImportNode(startPara, true) as Paragraph;
        Paragraph importedEnd = importer.ImportNode(endPara, true) as Paragraph;

        if (importedStart == null || importedEnd == null)
            throw new InvalidOperationException("Failed to import paragraphs.");

        newBody.AppendChild(importedStart);
        newBody.AppendChild(importedEnd);

        // ------------------------------------------------------------
        // 5. Validate that the styling is retained in the cloned paragraphs.
        // ------------------------------------------------------------
        Paragraph clonedFirst = extractedDoc.FirstSection.Body.Paragraphs[0];
        Paragraph clonedSecond = extractedDoc.FirstSection.Body.Paragraphs[1];

        if (clonedFirst == null || clonedSecond == null)
            throw new InvalidOperationException("Cloned paragraphs not found.");

        Run runFirst = clonedFirst.Runs.Count > 0 ? clonedFirst.Runs[0] : null;
        Run runSecond = clonedSecond.Runs.Count > 0 ? clonedSecond.Runs[0] : null;

        if (runFirst == null || runSecond == null)
            throw new InvalidOperationException("Runs missing in cloned paragraphs.");

        bool firstBold = runFirst.Font.Bold;
        bool secondItalic = runSecond.Font.Italic;

        if (!firstBold || !secondItalic)
            throw new InvalidOperationException("Extracted paragraph styling does not match the source.");

        // ------------------------------------------------------------
        // 6. Save the extracted document and a JSON test result.
        // ------------------------------------------------------------
        extractedDoc.Save("extracted.docx");

        var testResult = new
        {
            Test = "ParagraphStylingExtraction",
            Passed = true,
            Details = new
            {
                FirstParagraphBold = firstBold,
                SecondParagraphItalic = secondItalic
            }
        };

        File.WriteAllText("test-result.json",
            JsonConvert.SerializeObject(testResult, Formatting.Indented));

        Console.WriteLine("Extraction test passed. Results written to extracted.docx and test-result.json.");
    }
}
