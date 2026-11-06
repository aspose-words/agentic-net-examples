using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with some text.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world!");

        // Create the revised document where the word "world" is removed.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello!");

        // Set compare options (no ShowDeletedContent property in this version).
        CompareOptions options = new CompareOptions();

        // Perform the comparison.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Verify that at least one revision (deletion) exists.
        if (original.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected at least one revision after comparison.");
        }

        // Save the comparison result.
        string outputPath = "comparison_show_deleted.docx";
        original.Save(outputPath);
    }
}
