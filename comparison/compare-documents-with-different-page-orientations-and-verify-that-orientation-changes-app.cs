using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the original document with default (portrait) orientation.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This document is in portrait orientation.");
        // Ensure the orientation is explicitly set to Portrait for clarity.
        builderOriginal.PageSetup.Orientation = Orientation.Portrait;

        // Create the revised document and change its orientation to Landscape.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.PageSetup.Orientation = Orientation.Landscape;
        builderRevised.Writeln("This document is in landscape orientation.");

        // Compare the two documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision exists.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Check that a format change revision (orientation change) is present.
        bool hasFormatChange = false;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
            {
                hasFormatChange = true;
                break;
            }
        }

        if (!hasFormatChange)
            throw new InvalidOperationException("Expected a format change revision for orientation, but none was found.");

        // Save the compared document with revisions.
        string outputPath = "OrientationComparison.docx";
        original.Save(outputPath);
    }
}
