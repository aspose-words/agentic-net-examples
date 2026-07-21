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
        builderOriginal.Writeln("This is a sample paragraph.");
        builderOriginal.PageSetup.Orientation = Orientation.Portrait;

        // Create the revised document with landscape orientation.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("This is a sample paragraph.");
        builderRevised.PageSetup.Orientation = Orientation.Landscape;

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision exists.
        int totalRevisions = original.Revisions.Count;
        Console.WriteLine($"Total revisions after comparison: {totalRevisions}");

        // Look for a format change revision that affects the section's page orientation.
        bool orientationRevisionFound = false;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange &&
                rev.ParentNode != null &&
                rev.ParentNode.NodeType == NodeType.Section)
            {
                orientationRevisionFound = true;
                Console.WriteLine("Orientation change detected as a format revision on a section node.");
                break;
            }
        }

        if (!orientationRevisionFound)
        {
            Console.WriteLine("No orientation revision was detected.");
        }

        // Save the document that now contains the revisions.
        original.Save("OrientationComparison.docx");
    }
}
