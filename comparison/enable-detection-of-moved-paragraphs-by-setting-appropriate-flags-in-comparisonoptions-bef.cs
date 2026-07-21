using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DetectMovedParagraphs
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1.");
        builder.Writeln("Paragraph 2."); // This paragraph will be moved.
        builder.Writeln("Paragraph 3.");

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);

        // Move the second paragraph to the end (after paragraph 3).
        Paragraph paragraphToMove = revised.FirstSection.Body.Paragraphs[1]; // Index 1 = "Paragraph 2."
        Node referenceNode = revised.FirstSection.Body.Paragraphs[2]; // "Paragraph 3."
        revised.FirstSection.Body.InsertAfter(paragraphToMove, referenceNode);

        // Configure comparison options to detect moved paragraphs.
        CompareOptions compareOptions = new CompareOptions
        {
            CompareMoves = true,                     // Enable move detection.
            IgnoreFormatting = false,
            IgnoreCaseChanges = false,
            IgnoreComments = false,
            IgnoreTables = false,
            IgnoreFields = false,
            IgnoreFootnotes = false,
            IgnoreTextboxes = false,
            IgnoreHeadersAndFooters = false,
            Target = ComparisonTargetType.New      // Use the revised document as the target.
        };

        // Perform the comparison.
        original.Compare(revised, "DemoAuthor", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(outputDir, "ComparisonResult.docx");
        original.Save(resultPath);

        // Inspect revisions to find moved paragraphs.
        int moveFromCount = 0;
        int moveToCount = 0;
        foreach (Paragraph para in original.FirstSection.Body.Paragraphs)
        {
            if (para.IsMoveFromRevision) moveFromCount++;
            if (para.IsMoveToRevision) moveToCount++;
        }

        // Output detection summary.
        Console.WriteLine($"Moved paragraph 'from' revisions detected: {moveFromCount}");
        Console.WriteLine($"Moved paragraph 'to' revisions detected: {moveToCount}");
        Console.WriteLine($"Comparison result saved to: {resultPath}");
    }
}
