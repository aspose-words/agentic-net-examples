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
        DocumentBuilder builder = new DocumentBuilder(original);

        // Section 1
        builder.Writeln("Section 1 - Original content.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2
        builder.Writeln("Section 2 - Original content.");

        // Save the original for reference (optional).
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        // Modify text in Section 1.
        revBuilder.MoveToSection(0);
        revBuilder.Writeln("Section 1 - Revised content.");

        // Modify text in Section 2.
        revBuilder.MoveToSection(1);
        revBuilder.Writeln("Section 2 - Revised content.");

        // Save the revised document (optional).
        string revisedPath = Path.Combine(Directory.GetCurrentDirectory(), "Revised.docx");
        revised.Save(revisedPath);

        // Set up compare options – we will use the default options but specify the target document.
        CompareOptions compareOptions = new CompareOptions
        {
            Target = ComparisonTargetType.New // Use the revised document as the target during comparison.
        };

        // Perform the comparison. Revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "Compared.docx");
        original.Save(resultPath);

        // Analyze revisions only in Section 2 (index 1).
        int sectionIndexToInspect = 1;
        int revisionsInSection = 0;

        foreach (Revision rev in original.Revisions)
        {
            // Get the section that contains the revision's parent node.
            Node? parent = rev.ParentNode;
            if (parent == null) continue;

            Section? section = parent.GetAncestor(NodeType.Section) as Section;
            if (section == null) continue;

            // Sections are stored in the document's Sections collection.
            int currentIndex = original.Sections.IndexOf(section);
            if (currentIndex == sectionIndexToInspect)
            {
                revisionsInSection++;
                Console.WriteLine($"Revision in Section {sectionIndexToInspect + 1}: Type={rev.RevisionType}, Text=\"{parent.GetText().Trim()}\"");
            }
        }

        Console.WriteLine($"Total revisions found in Section {sectionIndexToInspect + 1}: {revisionsInSection}");
    }
}
