using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the first document (original) with some content.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is the original document.");
        builderOriginal.Writeln("It contains a single paragraph.");

        // Save the original for reference (optional).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Create the second document (revised) with different content.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is the revised document.");
        builderRevised.Writeln("It now contains two paragraphs.");
        builderRevised.Writeln("Additional line added.");

        // Configure compare options to set the comparison target to the new document.
        CompareOptions compareOptions = new CompareOptions
        {
            // When set to New, the document passed as the argument to Compare becomes the base.
            // Revisions will therefore appear in the document on which Compare is called (revised).
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be recorded in the 'revised' document.
        revised.Compare(original, "Comparer", DateTime.Now, compareOptions);

        // Verify that revisions were created.
        if (revised.Revisions.Count == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // Save the revised document which now contains the revisions.
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "RevisedWithRevisions.docx");
        revised.Save(revisedPath);
    }
}
