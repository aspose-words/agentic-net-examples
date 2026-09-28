using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
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
        builderRev.Writeln("First paragraph.");  // Original first paragraph.

        // Perform the comparison. In the current Aspose.Words version, moved paragraph detection
        // is enabled by default, so no additional CompareOptions are required.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Count all revisions produced by the comparison.
        int totalRevisions = original.Revisions.Count;

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MovedParagraphComparison.docx");
        original.Save(outputPath);

        // Output summary to the console.
        Console.WriteLine($"Total revisions detected: {totalRevisions}");
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
