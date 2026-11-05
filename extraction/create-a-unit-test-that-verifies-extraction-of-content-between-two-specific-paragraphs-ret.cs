using System;
using Aspose.Words;
using Aspose.Words.Fonts;

public class Program
{
    public static void Main()
    {
        // Create a source document with distinct styling on each paragraph.
        var source = new Document();
        var builder = new DocumentBuilder(source);

        builder.Writeln("Paragraph 1 - normal");

        builder.Font.Bold = true;
        builder.Writeln("Paragraph 2 - bold");

        builder.Font.Bold = false;
        builder.Font.Italic = true;
        builder.Writeln("Paragraph 3 - italic");

        builder.Font.Italic = false;
        builder.Font.Underline = Underline.Single;
        builder.Writeln("Paragraph 4 - underline");

        // Identify the start and end paragraphs (2nd and 3rd).
        Paragraph startPara = source.FirstSection.Body.Paragraphs[1];
        Paragraph endPara = source.FirstSection.Body.Paragraphs[2];

        if (startPara == null || endPara == null)
            throw new InvalidOperationException("Boundary paragraphs not found.");

        // Prepare the destination document.
        var extracted = new Document();
        extracted.RemoveAllChildren();

        var section = new Section(extracted);
        extracted.AppendChild(section);
        var body = new Body(extracted);
        section.AppendChild(body);

        // Use NodeImporter to copy paragraphs while preserving formatting.
        var importer = new NodeImporter(source, extracted, ImportFormatMode.KeepSourceFormatting);

        Paragraph importedStart = (Paragraph)importer.ImportNode(startPara, true);
        Paragraph importedEnd = (Paragraph)importer.ImportNode(endPara, true);

        body.AppendChild(importedStart);
        body.AppendChild(importedEnd);

        // Validate that exactly two paragraphs were extracted.
        if (extracted.FirstSection.Body.Paragraphs.Count != 2)
            throw new InvalidOperationException("Extracted paragraph count mismatch.");

        // Validate that the styling of each extracted paragraph matches the source.
        ValidateParagraphStyling(source, extracted, sourceIndex: 1, extractedIndex: 0);
        ValidateParagraphStyling(source, extracted, sourceIndex: 2, extractedIndex: 1);

        // Save the extracted document (optional verification artifact).
        extracted.Save("extracted.docx");

        Console.WriteLine("Extraction test passed.");
    }

    private static void ValidateParagraphStyling(Document source, Document extracted, int sourceIndex, int extractedIndex)
    {
        Paragraph srcPara = source.FirstSection.Body.Paragraphs[sourceIndex];
        Paragraph extPara = extracted.FirstSection.Body.Paragraphs[extractedIndex];

        if (srcPara == null || extPara == null)
            throw new InvalidOperationException("Paragraph for validation not found.");

        if (srcPara.Runs.Count == 0 || extPara.Runs.Count == 0)
            throw new InvalidOperationException("Paragraph runs missing.");

        Run srcRun = srcPara.Runs[0];
        Run extRun = extPara.Runs[0];

        bool stylingMatches = srcRun.Font.Bold == extRun.Font.Bold &&
                               srcRun.Font.Italic == extRun.Font.Italic &&
                               srcRun.Font.Underline == extRun.Font.Underline;

        if (!stylingMatches)
            throw new InvalidOperationException("Styling mismatch between source and extracted paragraph.");
    }
}
