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

        // Create the revised document with a difference.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Hello revised world.");

        // Set up compare options. Use the default target (Current) so revisions are added to the
        // document on which Compare is called (the original document).
        CompareOptions compareOptions = new CompareOptions
        {
            // No need to set Target; the default is ComparisonTargetType.Current.
        };

        // Perform the comparison. Revisions will be added to the 'original' document.
        original.Compare(revised, "John Doe", DateTime.Now, compareOptions);

        // Verify that revisions exist in the original document.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions in the original document, but none were found.");

        // Save both documents for inspection.
        original.Save("original_with_revisions.docx");
        revised.Save("revised.docx");

        // Output the number of revisions found in the original document.
        Console.WriteLine($"Revisions in original document: {original.Revisions.Count}");
    }
}
