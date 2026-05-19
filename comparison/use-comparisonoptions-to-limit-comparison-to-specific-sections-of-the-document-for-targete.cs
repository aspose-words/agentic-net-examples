using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with two sections.
        var original = new Document();
        var builder = new DocumentBuilder(original);
        builder.Writeln("Original Section 1 - Paragraph 1.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Original Section 2 - Paragraph 1.");

        // Create the revised document with changes in both sections.
        var revised = new Document();
        var revBuilder = new DocumentBuilder(revised);
        revBuilder.Writeln("Edited Section 1 - Paragraph 1."); // changed text in section 1
        revBuilder.InsertBreak(BreakType.SectionBreakNewPage);
        revBuilder.Writeln("Edited Section 2 - Paragraph 1."); // changed text in section 2

        // Set up compare options (default options are sufficient for this example).
        var compareOptions = new CompareOptions();

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the compared document.
        string outputDir = Directory.GetCurrentDirectory();
        string comparedPath = Path.Combine(outputDir, "compared.docx");
        original.Save(comparedPath);

        // Analyze revisions only in the second section (index 1).
        int targetSectionIndex = 1; // zero‑based index: 0 = first section, 1 = second section
        int revisionsInTargetSection = 0;

        foreach (Revision rev in original.Revisions)
        {
            // Guard against null ParentNode (e.g., style revisions).
            if (rev.ParentNode == null) continue;

            // Find the section that contains the revision's parent node.
            var section = rev.ParentNode.GetAncestor(NodeType.Section) as Section;
            if (section == null) continue;

            // Determine the section's index within the document.
            int sectionIndex = original.Sections.IndexOf(section);
            if (sectionIndex == targetSectionIndex)
                revisionsInTargetSection++;
        }

        // Write a simple report about the filtered revisions.
        string reportPath = Path.Combine(outputDir, "revision_report.txt");
        string reportContent = $"Revisions found in section {targetSectionIndex + 1}: {revisionsInTargetSection}";
        File.WriteAllText(reportPath, reportContent);
    }
}
