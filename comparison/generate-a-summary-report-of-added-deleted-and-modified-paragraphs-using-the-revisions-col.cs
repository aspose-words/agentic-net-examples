using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Comparing;

public class RevisionSummaryReport
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document.
        Document original = new Document();
        DocumentBuilder builderOrig = new DocumentBuilder(original);
        builderOrig.Writeln("Paragraph 1: This is the first paragraph.");
        builderOrig.Writeln("Paragraph 2: This is the second paragraph.");
        builderOrig.Writeln("Paragraph 3: This is the third paragraph.");
        original.Save(Path.Combine(outputDir, "Original.docx"));

        // Create the revised document with intentional changes.
        Document revised = new Document();
        DocumentBuilder builderRev = new DocumentBuilder(revised);
        // Change style of first paragraph to cause a format change revision.
        builderRev.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
        builderRev.Writeln("Paragraph 1: This is the first paragraph."); // same text, different style
        // Modify text of second paragraph.
        builderRev.Writeln("Paragraph 2: This paragraph has been modified.");
        // Delete third paragraph by simply not adding it.
        // Add a new fourth paragraph.
        builderRev.Writeln("Paragraph 4: This is a newly added paragraph.");
        revised.Save(Path.Combine(outputDir, "Revised.docx"));

        // Compare documents – revisions will be added to the original document.
        original.Compare(revised, "Comparer", DateTime.Now);

        // Verify that revisions were generated.
        if (original.Revisions.Count == 0)
            throw new InvalidOperationException("No revisions were detected after comparison.");

        // Summarize revisions by type.
        int addedCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Insertion);
        int deletedCount = original.Revisions.Count(r => r.RevisionType == RevisionType.Deletion);
        int modifiedCount = original.Revisions.Count(r => r.RevisionType == RevisionType.FormatChange);

        // Build a textual report.
        string report = $"Revision Summary Report{Environment.NewLine}" +
                        $"--------------------------------{Environment.NewLine}" +
                        $"Added paragraphs (Insertions)   : {addedCount}{Environment.NewLine}" +
                        $"Deleted paragraphs (Deletions) : {deletedCount}{Environment.NewLine}" +
                        $"Modified paragraphs (Format)   : {modifiedCount}{Environment.NewLine}";

        // Output to console.
        Console.WriteLine(report);

        // Save the report to a text file.
        File.WriteAllText(Path.Combine(outputDir, "RevisionSummary.txt"), report);

        // Save the compared document showing revisions.
        original.Save(Path.Combine(outputDir, "ComparisonResult.docx"));
    }
}
