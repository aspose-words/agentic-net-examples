using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class ComparisonOptionsSectionExample
{
    public static void Main()
    {
        // Prepare a folder for all artifacts.
        string artifactsDir = Path.Combine(Directory.GetCurrentDirectory(), "Artifacts");
        Directory.CreateDirectory(artifactsDir);

        // ---------- Create the original document ----------
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Section 1
        builder.Writeln("Section 1 - Original text.");
        // Insert a section break.
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Original text.");

        string originalPath = Path.Combine(artifactsDir, "Original.docx");
        original.Save(originalPath);

        // ---------- Create the revised document ----------
        // Clone the original and modify each section.
        Document revised = (Document)original.Clone(true);

        // Modify Section 1 text.
        Paragraph sec1Para = revised.FirstSection.Body.FirstParagraph;
        sec1Para.Runs[0].Text = "Section 1 - Revised text.";

        // Modify Section 2 text.
        Section secondSection = revised.Sections[1];
        Paragraph sec2Para = secondSection.Body.FirstParagraph;
        sec2Para.Runs[0].Text = "Section 2 - Revised text.";

        string revisedPath = Path.Combine(artifactsDir, "Revised.docx");
        revised.Save(revisedPath);

        // ---------- Set up compare options ----------
        CompareOptions compareOptions = new CompareOptions
        {
            // Example: ignore headers/footers (not present) to demonstrate option usage.
            IgnoreHeadersAndFooters = true,
            // Use the revised document as the target for comparison.
            Target = ComparisonTargetType.New
        };

        // ---------- Perform comparison ----------
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        string resultPath = Path.Combine(artifactsDir, "ComparisonResult.docx");
        original.Save(resultPath);

        // ---------- Analyze revisions limited to Section 1 ----------
        int revisionsInSection1 = 0;
        foreach (Revision rev in original.Revisions)
        {
            // Find the section that contains the revision's parent node.
            Node sectionNode = rev.ParentNode?.GetAncestor(NodeType.Section);
            if (sectionNode is Section sec && original.Sections.IndexOf(sec) == 0) // first section has index 0
            {
                revisionsInSection1++;
            }
        }

        // Output the count of revisions that belong to the first section.
        Console.WriteLine($"Revisions in Section 1: {revisionsInSection1}");
    }
}
