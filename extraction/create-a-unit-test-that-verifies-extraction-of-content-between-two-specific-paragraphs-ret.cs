using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;

public class ExtractionStyleTest
{
    public static void Main()
    {
        // Create a source document with styled paragraphs.
        Document sourceDoc = new Document();
        DocumentBuilder builder = new DocumentBuilder(sourceDoc);

        // Paragraph 1 – bold text.
        builder.Font.Bold = true;
        builder.Writeln("Bold paragraph");

        // Paragraph 2 – normal text.
        builder.Font.Bold = false;
        builder.Font.Italic = false;
        builder.Writeln("Normal paragraph");

        // Paragraph 3 – italic text.
        builder.Font.Italic = true;
        builder.Writeln("Italic paragraph");

        // Paragraph 4 – underline text.
        builder.Font.Italic = false;
        builder.Font.Underline = Underline.Single;
        builder.Writeln("Underlined paragraph");

        // Identify the start and end paragraphs for extraction (Paragraph 2 and 3).
        Paragraph startPara = sourceDoc.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = sourceDoc.FirstSection.Body.Paragraphs[2];
        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Create a new document to hold the extracted content.
        Document extractedDoc = new Document();
        extractedDoc.RemoveAllChildren();

        // Build a valid document structure (Section + Body).
        Section targetSection = new Section(extractedDoc);
        extractedDoc.AppendChild(targetSection);
        Body targetBody = new Body(extractedDoc);
        targetSection.AppendChild(targetBody);

        // Clone and import paragraphs from start to end (inclusive).
        int startIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(startPara);
        int endIndex = sourceDoc.FirstSection.Body.Paragraphs.IndexOf(endPara);
        for (int i = startIndex; i <= endIndex; i++)
        {
            Paragraph srcParagraph = sourceDoc.FirstSection.Body.Paragraphs[i];
            Paragraph importedParagraph = (Paragraph)extractedDoc.ImportNode(srcParagraph, true, ImportFormatMode.KeepSourceFormatting);
            targetBody.AppendChild(importedParagraph);
        }

        // Validate that the extracted document contains the expected number of paragraphs.
        if (extractedDoc.FirstSection.Body.Paragraphs.Count != 2)
            throw new InvalidOperationException("Extracted paragraph count mismatch.");

        // Verify that styling of each run is preserved.
        for (int i = 0; i < 2; i++)
        {
            Paragraph srcPara = sourceDoc.FirstSection.Body.Paragraphs[startIndex + i];
            Paragraph extPara = extractedDoc.FirstSection.Body.Paragraphs[i];

            if (srcPara.Runs.Count != extPara.Runs.Count)
                throw new InvalidOperationException($"Run count mismatch in paragraph {i + 1}.");

            for (int j = 0; j < srcPara.Runs.Count; j++)
            {
                Run srcRun = srcPara.Runs[j];
                Run extRun = extPara.Runs[j];

                if (srcRun.Font.Bold != extRun.Font.Bold ||
                    srcRun.Font.Italic != extRun.Font.Italic ||
                    srcRun.Font.Underline != extRun.Font.Underline)
                {
                    throw new InvalidOperationException($"Styling mismatch in paragraph {i + 1}, run {j + 1}.");
                }
            }
        }

        // Save the extracted document (optional verification artifact).
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "extracted.docx");
        extractedDoc.Save(outputPath);

        // Indicate success.
        Console.WriteLine("Extraction test passed. Output saved to: " + outputPath);
    }
}
