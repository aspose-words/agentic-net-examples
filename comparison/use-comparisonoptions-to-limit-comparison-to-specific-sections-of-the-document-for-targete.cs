using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonExample
{
    public static void Main()
    {
        // Create the original document with three sections.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        builder.Writeln("Section 1 - Original content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 - Original content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 3 - Original content.");

        // Save the original document (optional, for inspection).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Create the revised document where only Section 2 is changed.
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        revBuilder.Writeln("Section 1 - Original content.");
        revBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        revBuilder.Writeln("Section 2 - Revised content with modification.");
        revBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        revBuilder.Writeln("Section 3 - Original content.");

        // Set compare options to ignore formatting changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform the comparison; revisions are stored in the original document.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Target section index (zero‑based). We want revisions only from Section 2.
        int targetSectionIndex = 1;
        int revisionsInTargetSection = 0;

        foreach (Revision rev in original.Revisions)
        {
            // Walk up the node hierarchy to find the containing Section.
            Node? node = rev.ParentNode;
            while (node != null && !(node is Section))
            {
                node = node.ParentNode;
            }

            if (node is Section section && original.Sections.IndexOf(section) == targetSectionIndex)
            {
                revisionsInTargetSection++;
            }
        }

        // Report the findings.
        Console.WriteLine($"Total revisions detected: {original.Revisions.Count}");
        Console.WriteLine($"Revisions in Section {targetSectionIndex + 1}: {revisionsInTargetSection}");

        // Save the document that now contains the revisions.
        string comparedPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(comparedPath);
    }
}
