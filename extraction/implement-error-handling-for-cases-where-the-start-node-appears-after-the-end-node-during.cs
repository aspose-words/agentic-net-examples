using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Newtonsoft.Json;

public class ExtractionExample
{
    public static void Main()
    {
        // Ensure deterministic output folder (current directory)
        string inputPath = "sample.docx";
        string extractedPath = "extracted.docx";
        string reportPath = "extracted_report.json";
        string errorPath = "extraction_error.txt";

        // Create a sample document with identifiable paragraphs
        Document sampleDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sampleDoc);
        builder.Writeln("Intro paragraph.");
        builder.Writeln("Start marker paragraph.");
        builder.Writeln("Middle paragraph 1.");
        builder.Writeln("Middle paragraph 2.");
        builder.Writeln("End marker paragraph.");
        builder.Writeln("Conclusion paragraph.");
        sampleDoc.Save(inputPath);

        // Load the document for extraction
        Document loadedDoc = new Document(inputPath);
        Body body = loadedDoc.FirstSection.Body;

        // Locate start and end paragraphs by their text content
        Paragraph? startPara = FindParagraphByText(body, "Start marker paragraph.");
        Paragraph? endPara = FindParagraphByText(body, "End marker paragraph.");

        // Validate that both markers were found
        if (startPara == null || endPara == null)
        {
            File.WriteAllText(errorPath, "Start or end marker paragraph not found.");
            return;
        }

        // Determine the positions of the markers within the body
        int startIndex = body.Paragraphs.IndexOf(startPara);
        int endIndex = body.Paragraphs.IndexOf(endPara);

        // Error handling: start appears after end
        if (startIndex > endIndex)
        {
            File.WriteAllText(errorPath, $"Invalid range: start index ({startIndex}) is after end index ({endIndex}).");
            return;
        }

        try
        {
            // Extract paragraphs from start to end inclusive
            List<string> extractedTexts = new List<string>();
            Document resultDoc = new Document();
            resultDoc.RemoveAllChildren();
            Section resultSection = new Section(resultDoc);
            resultDoc.AppendChild(resultSection);
            Body resultBody = new Body(resultDoc);
            resultSection.AppendChild(resultBody);

            for (int i = startIndex; i <= endIndex; i++)
            {
                Paragraph srcPara = (Paragraph)body.Paragraphs[i];
                Paragraph clonedPara = (Paragraph)srcPara.Clone(true);
                resultBody.AppendChild(clonedPara);
                extractedTexts.Add(srcPara.GetText().TrimEnd('\r', '\n'));
            }

            // Save the extracted content document
            resultDoc.Save(extractedPath);

            // Serialize extracted texts to JSON report
            string jsonReport = JsonConvert.SerializeObject(
                new { ExtractedParagraphs = extractedTexts },
                Formatting.Indented);
            File.WriteAllText(reportPath, jsonReport);

            // Validate that output files were created
            if (!File.Exists(extractedPath) || !File.Exists(reportPath))
                throw new InvalidOperationException("Extraction output files were not created as expected.");
        }
        catch (Exception ex)
        {
            // Write exception details to error file
            File.WriteAllText(errorPath, $"Extraction failed: {ex.Message}");
        }
    }

    private static Paragraph? FindParagraphByText(Body body, string text)
    {
        foreach (Paragraph para in body.Paragraphs)
        {
            if (para.GetText().TrimEnd('\r', '\n') == text)
                return para;
        }
        return null;
    }
}
