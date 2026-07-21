using System;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Hello world.");

        // Create the revised document with a different sentence.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Configure compare options so that revisions are applied to the new (revised) document.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions will be stored in the 'revised' document.
        revised.Compare(original, "John Doe", DateTime.Now, compareOptions);

        // Verify that revisions were created in the revised document.
        if (revised.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions in the revised document, but none were found.");

        // Save both documents to the local file system.
        original.Save("original.docx");
        revised.Save("revised_with_revisions.docx");
    }
}
