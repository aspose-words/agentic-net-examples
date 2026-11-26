using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonShowDeletedExample
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This line will stay.");
        builderOriginal.Writeln("This line will be deleted.");
        builderOriginal.Writeln("Another unchanged line.");

        // Create the revised document where one line is removed and a new line is added.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This line will stay.");
        // The line "This line will be deleted." is omitted to create a deletion.
        builderRevised.Writeln("This line is newly added.");
        builderRevised.Writeln("Another unchanged line.");

        // Configure compare options. The ShowDeletedContent property does not exist in the current API,
        // but deleted content is retained by default in the comparison output.
        CompareOptions options = new CompareOptions();

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that revisions were created.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result.
        string outputPath = "compare-show-deleted.docx";
        original.Save(outputPath);
    }
}
