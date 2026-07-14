using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Paragraph 1");
        builderOriginal.Writeln("Paragraph 2");
        builderOriginal.Writeln("Paragraph 3");

        // Create the revised document where the second paragraph is moved after the third.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Paragraph 1");
        builderRevised.Writeln("Paragraph 3");
        builderRevised.Writeln("Paragraph 2"); // Moved paragraph.

        // Configure compare options (default options are sufficient for moved‑paragraph detection).
        CompareOptions compareOptions = new CompareOptions();

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Count revisions.
        int totalRevisions = original.Revisions.Count;

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "MovedParagraphsComparison.docx");
        original.Save(outputPath);

        // Output summary to the console.
        Console.WriteLine($"Total revisions detected: {totalRevisions}");
        Console.WriteLine($"Comparison document saved to: {outputPath}");
    }
}
