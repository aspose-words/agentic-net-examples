using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Comparing;

public class Program
{
    public static void Main()
    {
        // Prepare output directory.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document with two sections.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);

        // Section 1.
        builder.Writeln("Section 1 original text.");
        // Insert a header to demonstrate that it will be ignored.
        builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
        builder.Writeln("Header for Section 1");
        builder.MoveToDocumentEnd();

        // Section 2.
        builder.InsertBreak(BreakType.SectionBreakNewPage);
        builder.Writeln("Section 2 original text.");
        // Add a table in Section 2 (will be ignored by compare options).
        builder.StartTable();
        builder.InsertCell();
        builder.Write("Cell A1");
        builder.InsertCell();
        builder.Write("Cell B1");
        builder.EndTable();

        // Save the original for reference (optional).
        string originalPath = Path.Combine(outputDir, "Original.docx");
        original.Save(originalPath);

        // Clone the original to create the revised version.
        Document revised = (Document)original.Clone(true);
        DocumentBuilder revBuilder = new DocumentBuilder(revised);

        // Modify only the text in Section 1.
        revBuilder.MoveToSection(0);
        Paragraph firstParagraph = revised.FirstSection.Body.FirstParagraph;
        firstParagraph.Runs[0].Text = "Section 1 revised text.";

        // Save the revised document (optional).
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        revised.Save(revisedPath);

        // Set up compare options to focus on body text only.
        CompareOptions compareOptions = new CompareOptions
        {
            IgnoreHeadersAndFooters = true,
            IgnoreTables = true,
            IgnoreFootnotes = false,
            IgnoreComments = true,
            IgnoreTextboxes = true,
            IgnoreFields = true,
            IgnoreFormatting = true,
            IgnoreCaseChanges = false,
            CompareMoves = false,
            Target = ComparisonTargetType.New
        };

        // Perform the comparison.
        string author = "John Doe";
        DateTime compareDate = DateTime.Now;
        original.Compare(revised, author, compareDate, compareOptions);

        // Save the comparison result.
        string resultPath = Path.Combine(outputDir, "Compared.docx");
        original.Save(resultPath);

        // Report revisions that belong to the first section only.
        Console.WriteLine("Revisions detected in Section 1:");
        foreach (Revision rev in original.Revisions)
        {
            Section revSection = rev.ParentNode?.GetAncestor(NodeType.Section) as Section;
            if (revSection != null && original.Sections.IndexOf(revSection) == 0)
            {
                Console.WriteLine($"- Type: {rev.RevisionType}, Text: \"{rev.ParentNode?.GetText().Trim()}\"");
            }
        }

        // Summary output.
        Console.WriteLine($"Total revisions: {original.Revisions.Count}");
        Console.WriteLine($"Comparison result saved to: {resultPath}");
    }
}
