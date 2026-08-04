using System;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Layout;

public class Program
{
    public static void Main()
    {
        // Create the original document with three paragraphs.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Paragraph 1: This text will stay.");
        builderOriginal.Writeln("Paragraph 2: This text will be deleted.");
        builderOriginal.Writeln("Paragraph 3: This text will stay as well.");

        // Create the revised document where the second paragraph is omitted.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Paragraph 1: This text will stay.");
        // The second paragraph is intentionally skipped to simulate a deletion.
        builderRevised.Writeln("Paragraph 3: This text will stay as well.");

        // Set up compare options. No special flags are needed for this scenario.
        CompareOptions compareOptions = new CompareOptions
        {
            // Use the default target (Current) – the original document will receive revisions.
            Target = ComparisonTargetType.Current
        };

        // Perform the comparison. The original document will contain the revisions.
        original.Compare(revised, "DemoAuthor", DateTime.Now, compareOptions);

        // Ensure that at least one revision (the deletion) was created.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected at least one revision after comparison.");

        // Configure revision layout to show the original (deleted) text in the output.
        original.LayoutOptions.RevisionOptions.ShowOriginalRevision = true;
        original.LayoutOptions.RevisionOptions.ShowRevisionMarks = true;

        // Save the comparison result.
        string outputPath = "ComparisonResult.docx";
        original.Save(outputPath);
    }
}
