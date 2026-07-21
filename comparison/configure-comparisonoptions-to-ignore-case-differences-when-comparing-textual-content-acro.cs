using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with mixed‑case text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello World");

        // Create the revised document that differs only by case.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("hello world");

        // Configure comparison to ignore case changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreCaseChanges = true
        };

        // Perform the comparison.
        original.Compare(revised, "Author", DateTime.Now, compareOptions);

        // Because case differences are ignored, there should be no revisions.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException($"Expected zero revisions, but found {original.Revisions.Count}.");

        // Save the (unchanged) original document as the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
