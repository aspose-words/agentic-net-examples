using System;
using Aspose.Words;

public class OrientationComparisonExample
{
    public static void Main()
    {
        // Create the first document with portrait orientation.
        Document portraitDoc = new Document();
        DocumentBuilder portraitBuilder = new DocumentBuilder(portraitDoc);
        portraitBuilder.PageSetup.Orientation = Orientation.Portrait;
        portraitBuilder.Writeln("This document is in portrait orientation.");

        // Create the second document with landscape orientation.
        Document landscapeDoc = new Document();
        DocumentBuilder landscapeBuilder = new DocumentBuilder(landscapeDoc);
        landscapeBuilder.PageSetup.Orientation = Orientation.Landscape;
        landscapeBuilder.Writeln("This document is in landscape orientation.");

        // Compare the two documents.
        string author = "OrientationComparer";
        DateTime compareDate = DateTime.Now;
        portraitDoc.Compare(landscapeDoc, author, compareDate);

        // Verify that at least one revision exists.
        int totalRevisions = portraitDoc.Revisions.Count;
        if (totalRevisions == 0)
        {
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");
        }

        // Check whether any revision corresponds to a section (orientation) change.
        bool orientationRevisionFound = false;
        foreach (Revision rev in portraitDoc.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange && rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Section)
            {
                // The format change on a Section node indicates an orientation change.
                orientationRevisionFound = true;
                break;
            }
        }

        if (!orientationRevisionFound)
        {
            throw new InvalidOperationException("Orientation change was not detected as a revision.");
        }

        // Save the comparison result.
        string outputPath = "OrientationComparison.docx";
        portraitDoc.Save(outputPath);

        // Output a simple summary.
        Console.WriteLine($"Comparison completed. Total revisions: {totalRevisions}");
        Console.WriteLine($"Orientation change detected as revision: {orientationRevisionFound}");
        Console.WriteLine($"Result saved to: {outputPath}");
    }
}
