using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonTargetExample
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Configure compare options to use the *new* document as the comparison target.
        // This means the other document (original) is treated as the base,
        // and revisions will be recorded in the document on which Compare is called (revised).
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be recorded in the revised document.
        revised.Compare(original, "JD", DateTime.Now, compareOptions);

        // Verify that revisions exist in the revised document.
        if (revised.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions in the revised document, but none were found.");

        // Save both documents to the current directory.
        string outputDir = Directory.GetCurrentDirectory();
        original.Save(Path.Combine(outputDir, "Original.docx"));
        revised.Save(Path.Combine(outputDir, "Revised_With_Revisions.docx"));

        // Simple console summary.
        Console.WriteLine($"Revisions in revised document: {revised.Revisions.Count}");
    }
}
