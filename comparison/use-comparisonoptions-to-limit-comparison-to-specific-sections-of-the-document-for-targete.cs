using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;
using Aspose.Words.Tables;

public class Program
{
    public static void Main()
    {
        // Create the original document with two sections.
        Document original = new Document();
        DocumentBuilder builderOriginal = new DocumentBuilder(original);
        builderOriginal.Writeln("Section 1 - Original text.");
        builderOriginal.InsertBreak(BreakType.SectionBreakNewPage);
        builderOriginal.Writeln("Section 2 - Original text.");
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "original.docx");
        original.Save(originalPath);

        // Create the revised document with changes in both sections.
        Document revised = new Document();
        DocumentBuilder builderRevised = new DocumentBuilder(revised);
        builderRevised.Writeln("Section 1 - Revised text."); // Change in Section 1
        builderRevised.InsertBreak(BreakType.SectionBreakNewPage);
        builderRevised.Writeln("Section 2 - Revised text."); // Change in Section 2
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "revised.docx");
        revised.Save(revisedPath);

        // Set compare options (e.g., ignore formatting) – focus on textual changes.
        CompareOptions options = new CompareOptions
        {
            IgnoreFormatting = true
        };

        // Perform comparison; revisions are added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, options);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "comparison_result.docx");
        original.Save(resultPath);

        // Analyze revisions only in Section 1 (index 0).
        int targetSectionIndex = 0;
        int revisionsInTargetSection = 0;

        foreach (Revision revision in original.Revisions)
        {
            // Find the ancestor Section of the revision's parent node.
            Node? ancestor = revision.ParentNode?.GetAncestor(NodeType.Section);
            if (ancestor is Section section)
            {
                int sectionIndex = original.Sections.IndexOf(section);
                if (sectionIndex == targetSectionIndex)
                {
                    revisionsInTargetSection++;
                }
            }
        }

        // Output the count of revisions found in the targeted section.
        Console.WriteLine($"Revisions in Section {targetSectionIndex + 1}: {revisionsInTargetSection}");
    }
}
