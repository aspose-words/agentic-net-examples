using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Drawing;

public class OrientationComparison
{
    public static void Main()
    {
        // Create the original document (portrait orientation).
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This document is in portrait orientation.");
        // Ensure the page setup is portrait (default).
        builderOriginal.PageSetup.Orientation = Orientation.Portrait;

        // Create the revised document (landscape orientation).
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This document is in landscape orientation.");
        // Change the page orientation to landscape.
        builderRevised.PageSetup.Orientation = Orientation.Landscape;

        // Verify that both documents start without revisions.
        if (original.Revisions.Count != 0 || revised.Revisions.Count != 0)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will receive revisions.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Count total revisions and specifically format-change revisions (orientation change).
        int totalRevisions = original.Revisions.Count;
        int formatRevisions = original.Revisions.Count(r => r.RevisionType == RevisionType.FormatChange);

        // Verify that at least one format-change revision exists (orientation change).
        if (formatRevisions == 0)
            throw new InvalidOperationException("Expected a format-change revision for page orientation.");

        // Output revision summary to the console.
        Console.WriteLine($"Total revisions detected: {totalRevisions}");
        Console.WriteLine($"Format-change revisions (e.g., orientation): {formatRevisions}");

        // Save the compared document with revisions.
        string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "OrientationComparison.docx");
        original.Save(outputPath);
    }
}
