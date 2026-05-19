using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with the full text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");

        // Create the revised document where the word "world" is removed.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello ");

        // Configure comparison options.
        // Setting Target to ComparisonTargetType.New makes the revised document the base,
        // which keeps the deleted text visible in the comparison result.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Neither document contains revisions, satisfying the API requirement.
        original.Compare(revised, "Author", DateTime.Now, compareOptions);

        // Verify that at least one revision (the deletion) was generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(outputPath);
    }
}
