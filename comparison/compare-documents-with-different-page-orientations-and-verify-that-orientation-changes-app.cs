using System;
using Aspose.Words;
using Aspose.Words.Drawing;

public class Program
{
    public static void Main()
    {
        // Create the original document with portrait orientation (default).
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("This is a sample paragraph for orientation comparison.");

        // Create the revised document with landscape orientation.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.PageSetup.Orientation = Orientation.Landscape;
        builderRevised.Writeln("This is a sample paragraph for orientation comparison.");

        // Compare the documents. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that at least one revision exists.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("Expected revisions after comparison, but none were found.");

        // Verify that a format change revision (orientation change) is present.
        bool formatChangeFound = false;
        foreach (Revision rev in original.Revisions)
        {
            if (rev.RevisionType == RevisionType.FormatChange)
            {
                formatChangeFound = true;
                break;
            }
        }

        if (!formatChangeFound)
            throw new InvalidOperationException("Expected a format change revision for orientation, but none was found.");

        // Save the result document showing the revisions.
        string outputPath = "OrientationComparison.docx";
        original.Save(outputPath);

        // Output a simple confirmation.
        Console.WriteLine($"Comparison completed. Revisions count: {original.Revisions.Count}. Output saved to '{outputPath}'.");
    }
}
