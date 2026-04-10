using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Create the original document with two sections.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Section 1 content.
        builder.Writeln("Section 1 - Original paragraph.");
        builder.InsertBreak(BreakType.SectionBreakNewPage);

        // Section 2 content.
        builder.Writeln("Section 2 - Original paragraph.");

        // Save the original document.
        string originalPath = Path.Combine(Directory.GetCurrentDirectory(), "Original.docx");
        original.Save(originalPath);

        // Clone the original to create the edited version and modify both sections.
        Document edited = (Document)original.Clone(true);

        // Edit text in Section 1.
        Paragraph sec1Paragraph = edited.FirstSection.Body.FirstParagraph;
        sec1Paragraph.Runs[0].Text = "Section 1 - Edited paragraph.";

        // Edit text in Section 2.
        Section sec2 = edited.Sections[1];
        Paragraph sec2Paragraph = sec2.Body.FirstParagraph;
        sec2Paragraph.Runs[0].Text = "Section 2 - Edited paragraph.";

        // Save the edited document.
        string editedPath = Path.Combine(Directory.GetCurrentDirectory(), "Edited.docx");
        edited.Save(editedPath);

        // Configure comparison options to ignore formatting and comments (other elements are compared).
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreFormatting = true,
            IgnoreComments = true,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison. Revisions are added to the original document.
        original.Compare(edited, "Comparer", DateTime.Now, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(Directory.GetCurrentDirectory(), "ComparisonResult.docx");
        original.Save(resultPath);

        // Analyze revisions that belong only to Section 2 (index 1).
        int revisionsInSection2 = 0;
        foreach (Revision rev in original.Revisions)
        {
            Node? parent = rev.ParentNode;
            if (parent != null)
            {
                Section? parentSection = parent.GetAncestor(NodeType.Section) as Section;
                if (parentSection != null && original.Sections.IndexOf(parentSection) == 1)
                {
                    revisionsInSection2++;
                }
            }
        }

        // Output summary to the console.
        Console.WriteLine($"Total revisions detected: {original.Revisions.Count}");
        Console.WriteLine($"Revisions in Section 2: {revisionsInSection2}");
    }
}
