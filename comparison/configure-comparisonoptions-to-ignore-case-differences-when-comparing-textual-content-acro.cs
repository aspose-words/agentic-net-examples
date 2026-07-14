using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonIgnoreCaseExample
{
    public static void Main()
    {
        // Create the original document with mixed‑case text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello World!");

        // Create the revised document with the same text but different case.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("hello world!");

        // Configure comparison to ignore case changes.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreCaseChanges = true
        };

        // Perform the comparison.
        original.Compare(revised, "Author", DateTime.Now, compareOptions);

        // Verify that no revisions were created because case differences are ignored.
        if (original.Revisions.Count != 0)
            throw new InvalidOperationException("Case differences should have been ignored, but revisions were found.");

        // Save the result document.
        original.Save("ComparisonResult.docx");
    }
}
