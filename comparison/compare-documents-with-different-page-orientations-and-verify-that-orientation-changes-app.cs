using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the original document with portrait orientation.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Sample text for orientation comparison.");
        original.FirstSection.PageSetup.Orientation = Orientation.Portrait;

        // Create the revised document with landscape orientation.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Sample text for orientation comparison.");
        revised.FirstSection.PageSetup.Orientation = Orientation.Landscape;

        // Compare the two documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Ensure that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were generated after comparison.");

        // Look for a format‑change revision that originates from a Section node (orientation change).
        bool orientationRevisionFound = false;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange &&
                rev.ParentNode != null &&
                rev.ParentNode.NodeType == NodeType.Section)
            {
                orientationRevisionFound = true;
                break;
            }
        }

        if (!orientationRevisionFound)
            throw new InvalidOperationException("Orientation change was not detected as a revision.");

        // Save the document that now contains the revision tracking the orientation change.
        const string outputFile = "OrientationComparison.docx";
        original.Save(outputFile);
    }
}
