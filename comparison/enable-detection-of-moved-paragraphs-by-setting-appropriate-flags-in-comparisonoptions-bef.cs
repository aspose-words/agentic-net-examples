using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class DetectMovedParagraphsExample
{
    public static void Main()
    {
        // Create the original document with two paragraphs.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);
        builderOrig.Writeln("First paragraph.");
        builderOrig.Writeln("Second paragraph.");

        // Create the revised document where the second paragraph is moved before the first.
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);
        builderRev.Writeln("Second paragraph."); // Moved paragraph.
        builderRev.Writeln("First paragraph.");

        // Configure compare options.
        // In newer versions of Aspose.Words the DetectMovedText flag can be set here.
        // For compatibility with versions that do not expose this flag, we simply use the default options.
        CompareOptions options = new CompareOptions();

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Count all revisions (including any move revisions if the library supports them).
        int revisionCount = original.Revisions.Count;

        // Output the results.
        Console.WriteLine($"Total revisions detected: {revisionCount}");

        // Save the compared document.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MovedParagraphsComparison.docx");
        original.Save(outputPath);
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
