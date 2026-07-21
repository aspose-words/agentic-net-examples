using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Words;

public class RevisionSummaryExample
{
    public static void Main()
    {
        // Prepare output folder.
        string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
        Directory.CreateDirectory(outputDir);

        // Create the original document with a few paragraphs.
        Document original = new Document();
        DocumentBuilder builder = new DocumentBuilder(original);
        builder.Writeln("Paragraph 1: The quick brown fox.");
        builder.Writeln("Paragraph 2: Jumps over the lazy dog.");
        builder.Writeln("Paragraph 3: Lorem ipsum dolor sit amet.");
        builder.Writeln("Paragraph 4: Original content to be deleted.");
        string originalPath = Path.Combine(outputDir, "Original.docx");
        original.Save(originalPath);

        // Create the revised document with some changes:
        // - Delete paragraph 4.
        // - Modify paragraph 2.
        // - Add a new paragraph 5.
        Document revised = new Document();
        DocumentBuilder revBuilder = new DocumentBuilder(revised);
        revBuilder.Writeln("Paragraph 1: The quick brown fox."); // unchanged
        revBuilder.Writeln("Paragraph 2: Jumps over the energetic cat."); // modified
        revBuilder.Writeln("Paragraph 3: Lorem ipsum dolor sit amet."); // unchanged
        revBuilder.Writeln("Paragraph 5: This is a newly added paragraph."); // added
        string revisedPath = Path.Combine(outputDir, "Revised.docx");
        revised.Save(revisedPath);

        // Compare the documents. The original document will contain the revisions.
        original.Compare(revised, "Author", DateTime.Now);
        string comparedPath = Path.Combine(outputDir, "Compared.docx");
        original.Save(comparedPath);

        // Prepare collections for the summary.
        List<string> addedParagraphs = new();
        List<string> deletedParagraphs = new();
        List<string> modifiedParagraphs = new();

        // Build dictionaries of insertion and deletion revisions.
        Dictionary<string, string> deletionMap = new();
        Dictionary<string, string> insertionMap = new();

        foreach (Revision rev in original.Revisions)
        {
            if (rev.ParentNode?.NodeType != NodeType.Paragraph)
                continue;

            string text = rev.ParentNode.GetText().Trim();

            if (rev.RevisionType == RevisionType.Insertion)
                insertionMap[text] = text;
            else if (rev.RevisionType == RevisionType.Deletion)
                deletionMap[text] = text;
        }

        // Detect modifications by matching first words of deleted and inserted paragraphs.
        foreach (var del in deletionMap.Keys)
        {
            bool matched = false;
            foreach (var ins in insertionMap.Keys)
            {
                string[] delWords = del.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                string[] insWords = ins.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (delWords.Length > 0 && insWords.Length > 0 && delWords[0] == insWords[0])
                {
                    modifiedParagraphs.Add($"Original: \"{del}\" => Revised: \"{ins}\"");
                    matched = true;
                    break;
                }
            }
            if (!matched)
                deletedParagraphs.Add(del);
        }

        // Add remaining insertions that were not part of a modification.
        foreach (var ins in insertionMap.Keys)
        {
            bool alreadyHandled = false;
            foreach (var mod in modifiedParagraphs)
            {
                if (mod.Contains(ins))
                {
                    alreadyHandled = true;
                    break;
                }
            }
            if (!alreadyHandled)
                addedParagraphs.Add(ins);
        }

        // Build the summary report.
        var reportLines = new List<string>
        {
            "Revision Summary Report",
            "=======================",
            "",
            $"Total Added Paragraphs   : {addedParagraphs.Count}",
            $"Total Deleted Paragraphs : {deletedParagraphs.Count}",
            $"Total Modified Paragraphs: {modifiedParagraphs.Count}",
            ""
        };

        if (addedParagraphs.Count > 0)
        {
            reportLines.Add("Added Paragraphs:");
            foreach (string p in addedParagraphs)
                reportLines.Add("- " + p);
            reportLines.Add("");
        }

        if (deletedParagraphs.Count > 0)
        {
            reportLines.Add("Deleted Paragraphs:");
            foreach (string p in deletedParagraphs)
                reportLines.Add("- " + p);
            reportLines.Add("");
        }

        if (modifiedParagraphs.Count > 0)
        {
            reportLines.Add("Modified Paragraphs:");
            foreach (string p in modifiedParagraphs)
                reportLines.Add("- " + p);
            reportLines.Add("");
        }

        string reportPath = Path.Combine(outputDir, "RevisionSummary.txt");
        File.WriteAllLines(reportPath, reportLines);

        // Output locations for verification.
        Console.WriteLine("Original document saved to: " + originalPath);
        Console.WriteLine("Revised document saved to: " + revisedPath);
        Console.WriteLine("Compared document with revisions saved to: " + comparedPath);
        Console.WriteLine("Revision summary report saved to: " + reportPath);
    }
}
