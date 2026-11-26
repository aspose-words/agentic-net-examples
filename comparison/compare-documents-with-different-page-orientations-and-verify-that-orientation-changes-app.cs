using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with portrait orientation.
        Document docPortrait = new Document();
        DocumentBuilder builderPortrait = new DocumentBuilder(docPortrait);
        builderPortrait.PageSetup.Orientation = Orientation.Portrait;
        builderPortrait.Writeln("This is a sample paragraph for orientation comparison.");

        // Create the revised document with landscape orientation.
        Document docLandscape = new Document();
        DocumentBuilder builderLandscape = new DocumentBuilder(docLandscape);
        builderLandscape.PageSetup.Orientation = Orientation.Landscape;
        builderLandscape.Writeln("This is a sample paragraph for orientation comparison.");

        // Set compare options (default options are sufficient for detecting formatting changes).
        CompareOptions compareOptions = new CompareOptions();

        // Compare the two documents.
        docPortrait.Compare(docLandscape, "OrientationComparer", DateTime.Now, compareOptions);

        // Verify that the orientation change generated at least one revision.
        int revisionCount = docPortrait.Revisions.Count;
        if (revisionCount == 0)
        {
            throw new InvalidOperationException("No revisions detected; orientation change was not captured.");
        }

        // Output revision details.
        Console.WriteLine($"Total revisions detected: {revisionCount}");
        foreach (Revision rev in docPortrait.Revisions)
        {
            Console.WriteLine($"Revision Type: {rev.RevisionType}, Author: {rev.Author}");
        }

        // Save the comparison result.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OrientationComparisonResult.docx");
        docPortrait.Save(outputPath);
        Console.WriteLine($"Compared document saved to: {outputPath}");
    }
}
