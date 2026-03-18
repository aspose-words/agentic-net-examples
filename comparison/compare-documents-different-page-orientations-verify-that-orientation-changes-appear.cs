using System;
using Aspose.Words;
using Aspose.Words.Replacing;

class CompareOrientationRevisions
{
    static void Main()
    {
        // Create the original document.
        Document docOriginal = new Document();
        DocumentBuilder builder = new DocumentBuilder(docOriginal);
        builder.Writeln("This document will be used for orientation comparison.");

        // Ensure the original document is in Portrait orientation (default).
        docOriginal.FirstSection.PageSetup.Orientation = Orientation.Portrait;

        // Clone the original document to create the edited version.
        Document docEdited = (Document)docOriginal.Clone(true);

        // Change the orientation of the edited document to Landscape.
        docEdited.FirstSection.PageSetup.Orientation = Orientation.Landscape;

        // Verify that both documents have no revisions before comparison.
        if (docOriginal.HasRevisions || docEdited.HasRevisions)
            throw new InvalidOperationException("Documents must not contain revisions before comparison.");

        // Compare the documents. The original document will receive revisions that describe the differences.
        docOriginal.Compare(docEdited, "Author", DateTime.Now);

        // Iterate through the revisions to find the format change caused by orientation alteration.
        bool orientationRevisionFound = false;
        foreach (Revision rev in docOriginal.Revisions)
        {
            // Format changes are represented by RevisionType.FormatChange.
            if (rev.RevisionType == RevisionType.FormatChange)
            {
                // The parent node of a format revision is the node whose formatting changed.
                // In this case we expect it to be a Section node.
                if (rev.ParentNode != null && rev.ParentNode.NodeType == NodeType.Section)
                {
                    // Verify that the orientation stored in the revision matches the edited document's orientation.
                    Section revisedSection = (Section)rev.ParentNode;
                    if (revisedSection.PageSetup.Orientation == Orientation.Landscape)
                    {
                        orientationRevisionFound = true;
                        Console.WriteLine("Orientation change detected as a format revision.");
                    }
                }
            }
        }

        if (!orientationRevisionFound)
            Console.WriteLine("No orientation revision was found.");

        // Optional: Save the resulting document with revisions for manual inspection.
        docOriginal.Save("OrientationComparisonResult.docx");
    }
}
