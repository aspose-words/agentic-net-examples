using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with mixed‑case text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello World.");

        // Create the revised document with the same text but different case.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("hello world.");

        // Configure comparison options to ignore case changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreCaseChanges = true
        };

        // Perform the comparison. The author name and date are required.
        original.Compare(revised, "Author", DateTime.Now, compareOptions);

        // Because case differences are ignored, there should be no revisions.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Revisions were generated despite ignoring case changes.");

        // Save the result document (it will be identical to the original content).
        original.Save("ComparisonResult.docx");
    }
}
